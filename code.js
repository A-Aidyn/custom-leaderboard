function calculateAllRatings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('Raw Data');
  const lbSheet = ss.getSheetByName('Leaderboard');
  
  // Get all raw data (includes headers at index 0)
  const rawData = rawSheet.getDataRange().getValues();
  
  // Verify data exists
  if (rawData.length < 2) {
    SpreadsheetApp.getUi().alert('No data found in Raw Data sheet');
    return;
  }
  
  // Player state storage: { playerName: { rating: 1500, matches: 0 } }
  const players = {};
  
  // Group data by MatchID for efficient processing
  const matches = {};
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    const matchID = row[1];
    
    if (!matches[matchID]) matches[matchID] = [];
    matches[matchID].push({
      rowIndex: i,
      matchID: row[1],
      player: row[4],
      team: row[5],
      roundsWon: Number(row[6]) || 0,
      roundsLost: Number(row[7]) || 0,
      acs: Number(row[8]) || 0,
      kills: Number(row[9]) || 0,
      deaths: Number(row[10]) || 0,
      assists: Number(row[11]) || 0,
    });
  }
  
  // Process matches in sorted order (chronological)
  const sortedMatchIDs = Object.keys(matches).sort((a, b) => Number(a) - Number(b));
  
  Logger.log(`num of matches ${sortedMatchIDs.length}`);

  for (const matchID of sortedMatchIDs) {
    const matchPlayers = matches[matchID];
    
    // Need exactly 10 players (5v5)
    if (matchPlayers.length !== 10) {
      Logger.log(`Warning: Match ${matchID} has ${matchPlayers.length} players, expected 10`);
      continue;
    }
    
    // Split teams
    const teamA = matchPlayers.filter(p => p.team === 'A');
    const teamB = matchPlayers.filter(p => p.team === 'B');
    
    if (teamA.length !== 5 || teamB.length !== 5) {
      Logger.log(`Warning: Match ${matchID} has unbalanced teams`);
      continue;
    }
    
    // Initialize new players
    matchPlayers.forEach(p => {
      if (!players[p.player]) {
        players[p.player] = { rating: 1500, matches: 0 };
      }
    });
    
    // Calculate lobby averages
    const totalACS = matchPlayers.reduce((sum, p) => sum + p.acs, 0);
    const lobbyACS = totalACS / 10;
    
    const matchKDAs = matchPlayers.map(p => {
      return (p.kills + p.assists * 0.5) / Math.max(p.deaths, 1);
    });
    const lobbyKDA = matchKDAs.reduce((a, b) => a + b, 0) / 10;
    
    // Get team ratings
    const teamARatings = teamA.map(p => players[p.player].rating);
    const teamBRatings = teamB.map(p => players[p.player].rating);
    const teamAAvg = teamARatings.reduce((a, b) => a + b, 0) / 5;
    const teamBAvg = teamBRatings.reduce((a, b) => a + b, 0) / 5;
    
    // Team performance based on round margin
    const marginA = teamA[0].roundsWon - teamA[0].roundsLost;
    const teamAPerf = 0.5 + 0.5 * Math.tanh(marginA / 4);
    const teamBPerf = 1 - teamAPerf;
    
    const processPlayer = (p, teamPerf, myTeamAvg, oppAvg) => {
      const state = players[p.player];
      
      // Individual performance metrics
      const myKDA = (p.kills + p.assists * 0.5) / Math.max(p.deaths, 1);
      const acsRatio = lobbyACS > 0 ? p.acs / lobbyACS : 1;
      const kdaRatio = lobbyKDA > 0 ? Math.min(myKDA / lobbyKDA, 2.5) : 1;
      const perfIndex = 0.6 * acsRatio + 0.4 * kdaRatio;
      
      // Elo calculation
      const expected = 1 / (1 + Math.pow(10, (oppAvg - myTeamAvg) / 400));
      const baseChange = teamPerf - expected;
      
      // ASYMMETRIC PERFORMANCE MODIFIER (THE FIX)
      // Wins: High PI = multiply gain (good)
      // Losses: High PI = reduce loss (multiply by 2-PI, minimum 0.5x)
      const isGain = baseChange > 0;
      const perfMod = isGain ? perfIndex : Math.max(0.5, 2 - perfIndex);
      
      const kFactor = 32 * Math.max(1 - state.matches / 30, 0.5);
      const ratingChange = kFactor * baseChange * perfMod;
      
      // Update
      state.rating += ratingChange;
      state.matches++;

      // Debug logging (optional)
      Logger.log(`${p.player}: ${teamPerf.toFixed(3)} vs ${expected.toFixed(3)} exp, ` +
                  `PI=${perfIndex.toFixed(3)}, Δ=${ratingChange.toFixed(2)}, ` +
                  `new=${state.rating.toFixed(2)}`);
    };
    
    // Update all players
    teamA.forEach(p => processPlayer(p, teamAPerf, teamAAvg, teamBAvg));
    teamB.forEach(p => processPlayer(p, teamBPerf, teamBAvg, teamAAvg));
  }
  
  // Prepare Leaderboard data
  const lbData = Object.entries(players)
    .map(([name, data]) => ({
      player: name,
      rating: Math.round(data.rating),
      matches: data.matches
    }))
    .sort((a, b) => b.rating - a.rating)
    .map((p, index) => [index + 1, p.player, p.rating, p.matches]);


  // Write Leaderboard
  lbSheet.clear();
  lbSheet.getRange(1, 1, 1, 4).setValues([['Rank', 'Player', 'Rating', 'Matches']]);
  
  if (lbData.length > 0) {
    lbSheet.getRange(2, 1, lbData.length, 4).setValues(lbData);
  }
  
  // Formatting
  lbSheet.autoResizeColumns(1, 4);
  
  // Add conditional formatting for tiers (optional)
  const lastRow = lbData.length + 1;
  const ratingRange = lbSheet.getRange(2, 3, lbData.length, 1); // Column C (Rating)
  
  // Clear old rules
  lbSheet.setConditionalFormatRules([]);
  
  // Add color coding by rating tiers
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(2000)
      .setBackground('#FF4655') // Radiant - Red
      .setFontColor('#FFFFFF')
      .setRanges([ratingRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1800, 1999)
      .setBackground('#B784F7') // Immortal - Purple
      .setFontColor('#FFFFFF')
      .setRanges([ratingRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1600, 1799)
      .setBackground('#00B4D8') // Diamond - Blue
      .setFontColor('#FFFFFF')
      .setRanges([ratingRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1400, 1599)
      .setBackground('#A0A0A0') // Platinum - Gray
      .setFontColor('#FFFFFF')
      .setRanges([ratingRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1200, 1399)
      .setBackground('#FFD700') // Gold - Gold
      .setFontColor('#000000')
      .setRanges([ratingRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1000, 1199)
      .setBackground('#C0C0C0') // Silver - Silver
      .setFontColor('#000000')
      .setRanges([ratingRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(1000)
      .setBackground('#CD7F32') // Bronze - Bronze
      .setFontColor('#FFFFFF')
      .setRanges([ratingRange])
      .build()
  ];
  
  lbSheet.setConditionalFormatRules(rules);
  
  // Alert completion
  const matchCount = sortedMatchIDs.length;
  SpreadsheetApp.getUi().alert(
    'Ratings Updated', 
    `Processed ${lbData.length} players across ${matchCount} matches.`, 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏆 Leaderboard')
    .addItem('Calculate Ratings', 'calculateAllRatings')
    .addToUi();
}