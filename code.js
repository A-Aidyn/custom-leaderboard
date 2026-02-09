function validateMatch(matchPlayers) {
  for(const matchPlayer of matchPlayers) {
      if (matchPlayer.team !== 'A' && matchPlayer.team !== 'B') {
        Logger.log(`Warning: Match ${matchPlayer.matchID} has wrong format for team`)
        return false;  
      }
  }

  return true;
}

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
  
  // Uncertainty constants (Glicko-style RD, in rating points)
  const U_MIN = 50;    // floor: very confident active player
  const U_INIT = 200;  // new player starts highly uncertain
  const U_MAX = 200;   // cap so volatility can't explode
  const C = 20;        // idle growth rate: points per sqrt(day)
  const DECAY = 0.85;  // per-match shrink factor
  const K_MULT_MAX = 2.0;  // max K-factor boost from uncertainty
  
  // Player state storage: { playerName: { rating, matches, u, lastDate } }
  const players = {};
  
  // Per-row rating changes keyed by raw data row index
  const rowRatingChanges = {};
  
  // Group data by MatchID for efficient processing
  const matches = {};
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    const matchID = row[1];
    
    if (!matches[matchID]) matches[matchID] = [];
    matches[matchID].push({
      rowIndex: i,
      matchID: row[1],
      date: row[0],
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
        
    if (!validateMatch(matchPlayers)) {
      continue;
    }

    // Split teams
    const teamA = matchPlayers.filter(p => p.team === 'A');
    const teamB = matchPlayers.filter(p => p.team === 'B');
    
    if (teamA.length !== 5 || teamB.length !== 5) {
      Logger.log(`Warning: Match ${matchID} has incorrect number of players in each team. Team A: ${teamA.length}, Team B: ${teamB.length}`);
      continue;
    }
    
    // Validate match date (Google Sheets returns native Date objects)
    const matchDate = matchPlayers[0].date;
    if (!(matchDate instanceof Date) || isNaN(matchDate.getTime())) {
      Logger.log(`Error: Match ${matchID} has invalid date "${matchPlayers[0].date}", skipping`);
      continue;
    }
    
    // Initialize new players
    matchPlayers.forEach(p => {
      if (!players[p.player]) {
        players[p.player] = { rating: 1500, matches: 0, uncertainty: U_INIT, lastDate: null };
      }
    });
    
    // Grow uncertainty based on idle time since last match
    matchPlayers.forEach(p => {
      const state = players[p.player];
      if (state.lastDate) {
        var daysIdle = (matchDate - state.lastDate) / (1000 * 60 * 60 * 24);  // daysIdle range: [0, ...)
        if (daysIdle > 0) {
          // Glicko-style growth: u = min(U_MAX, sqrt(u^2 + C^2 * daysIdle))
          state.uncertainty = Math.min(U_MAX, Math.sqrt(state.uncertainty * state.uncertainty + C * C * daysIdle));
        }
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
    const teamAAvgRating = teamARatings.reduce((a, b) => a + b, 0) / 5;
    const teamBAvgRating = teamBRatings.reduce((a, b) => a + b, 0) / 5;
    
    // Team performance based on round margin
    const marginA = teamA[0].roundsWon - teamA[0].roundsLost;  // marginA range: [-13, 13] / {1} -> match difference cannot be 1 (in rare cases it could be 0)
    const teamAPerf = 0.5 + 0.5 * Math.tanh(marginA / 4);  // teamAPerf range: (0, 1)
    const teamBPerf = 1 - teamAPerf;
    
    const processPlayer = (
      p, 
      myTeamPerf, 
      myTeamAvgRating, 
      oppTeamAvgRating,
    ) => {
      const state = players[p.player];
      
      // Individual performance metrics
      const myKDA = (p.kills + p.assists * 0.5) / Math.max(p.deaths, 1);  // myKDA range: [0, ...]
      const acsRatio = lobbyACS > 0 ? p.acs / lobbyACS : 1;  // acsRatio range: [0, 10]
      const kdaRatio = lobbyKDA > 0 ? Math.min(myKDA / lobbyKDA, 2.5) : 1;  // kdaRatio range: [0, 2.5]
      const perfIndex = 0.6 * acsRatio + 0.4 * kdaRatio;  // perfIndex range: [0, 7]
      
      // Elo calculation
      const expected = 1 / (1 + Math.pow(10, (oppTeamAvgRating - myTeamAvgRating) / 400));  // expected range: (0, 1)
      const baseChange = myTeamPerf - expected;  // baseChange range: (-1, 1)
      
      // ASYMMETRIC PERFORMANCE MODIFIER
      // Wins: High perfIndex = multiply gain (good)
      // Losses: High perfIndex = reduce loss (multiply by 2-perfIndex, minimum 0.5x)
      const isGain = baseChange > 0;

      // isGain = true -> perfMod range: [0, 7]
      // isGain = false -> perfMod range: [0.5, 2]
      const perfMod = isGain ? perfIndex : Math.max(0.5, 2 - perfIndex);
      
      const baseK = 32 * Math.max(1 - state.matches / 30, 0.5);  // baseK range: [16, 32]
      
      // Uncertainty -> K multiplier: map u in [U_MIN, U_MAX] to kMult in [1, K_MULT_MAX]
      var u01 = Math.max(0, Math.min(1, (state.uncertainty - U_MIN) / (U_MAX - U_MIN)));  // u01 range: [0, 1]
      var kMult = 1 + u01 * (K_MULT_MAX - 1);  // kMult range: [1, 2]
      var kFactor = baseK * kMult;  // kFactor range: [16, 64]
      
      const ratingChange = kFactor * baseChange * perfMod;
      
      // Update rating, match count, decay uncertainty, record date
      state.rating += ratingChange;
      state.matches++;
      state.uncertainty = Math.max(U_MIN, state.uncertainty * DECAY);
      state.lastDate = matchDate;
      
      // Record rating change for this raw data row
      rowRatingChanges[p.rowIndex] = Math.round(ratingChange);

      // Debug logging (optional)
      Logger.log(`${p.player}: ${myTeamPerf.toFixed(3)} vs ${expected.toFixed(3)} exp, ` +
                  `PI=${perfIndex.toFixed(3)}, uncertainty=${state.uncertainty.toFixed(1)}, Δ=${ratingChange.toFixed(2)}, ` +
                  `new=${state.rating.toFixed(2)}`);
    };
    
    // Update all players
    teamA.forEach(p => processPlayer(p, teamAPerf, teamAAvgRating, teamBAvgRating));
    teamB.forEach(p => processPlayer(p, teamBPerf, teamBAvgRating, teamAAvgRating));
  }
  
  // --- Write Rating Change sheet (Raw Data columns + Rating Change) ---
  var rcSheet = ss.getSheetByName('Rating Change');
  if (!rcSheet) rcSheet = ss.insertSheet('Rating Change');
  rcSheet.clear();
  
  const rawHeader = rawData[0];
  const rcHeader = rawHeader.concat(['Rating Change']);
  rcSheet.getRange(1, 1, 1, rcHeader.length).setValues([rcHeader]);
  
  const rcRows = [];
  for (var i = 1; i < rawData.length; i++) {
    var change = rowRatingChanges[i] !== undefined ? rowRatingChanges[i] : '';
    rcRows.push(rawData[i].concat([change]));
  }
  // Sort: date asc, matchID asc, team asc, ACS desc, kills desc
  rcRows.sort(function(a, b) {
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;

    if (a[1].localeCompare(b[1]) === -1) return -1;
    if (a[1].localeCompare(b[1]) === 1) return 1;

    if (Number(a[6]) < Number(b[6])) return -1;
    if (Number(a[6]) > Number(b[6])) return 1;

    if (a[5] < b[5]) return -1;
    if (a[5] > b[5]) return 1;
    if (Number(b[8]) !== Number(a[8])) return Number(b[8]) - Number(a[8]);
    return Number(b[9]) - Number(a[9]);
  });
  
  for (let i = 0; i < 15; i ++) {
    Logger.log(`i: ${i} | rcRow[i]: ${rcRows[i]}`);
  }

  if (rcRows.length > 0) {
    rcSheet.getRange(2, 1, rcRows.length, rcHeader.length).setValues(rcRows);
  }
  rcSheet.autoResizeColumns(1, rcHeader.length);
  
  if (rcRows.length > 0) {
    // Format date column (column 1) as YYYY-MM-DD
    rcSheet.getRange(2, 1, rcRows.length, 1).setNumberFormat('yyyy-mm-dd');
    
    // Alternating background on matchID (col 2) and map (col 3) to separate matches
    var colors = ['#FFFFFF', '#E8EAF6'];
    var colorIdx = 0;
    var prevMatchID = rcRows[0][1];
    for (var r = 0; r < rcRows.length; r++) {
      if (rcRows[r][1] !== prevMatchID) {
        colorIdx = 1 - colorIdx;
        prevMatchID = rcRows[r][1];
      }
      rcSheet.getRange(r + 2, 1, 1, rcHeader.length - 1).setBackground(colors[colorIdx]);
    }
    
    // Color the Rating Change column: green for positive, red for negative
    var rcCol = rcHeader.length;
    var rcRange = rcSheet.getRange(2, rcCol, rcRows.length, 1);
    rcSheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground('#D9EAD3')
        .setFontColor('#006100')
        .setRanges([rcRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground('#F4CCCC')
        .setFontColor('#CC0000')
        .setRanges([rcRange])
        .build()
    ]);
  }
  
  // --- Write Leaderboard sheet ---
  const sortedPlayers = Object.entries(players)
    .map(([name, data]) => ({
      player: name,
      rating: Math.round(data.rating),
      matches: data.matches,
      uncertainty: Math.round(data.uncertainty),
      lastPlayed: data.lastDate
    }))
    .sort((a, b) => b.rating - a.rating);
  
  const lbHeader = ['Rank', 'Player', 'Rating', 'Matches', 'Uncertainty', 'Last Played'];
  const numCols = lbHeader.length;
  
  const lbData = sortedPlayers.map((p, index) => {
    return [index + 1, p.player, p.rating, p.matches, p.uncertainty, p.lastPlayed || ''];
  });

  lbSheet.clear();
  lbSheet.getRange(1, 1, 1, numCols).setValues([lbHeader]);
  
  if (lbData.length > 0) {
    lbSheet.getRange(2, 1, lbData.length, numCols).setValues(lbData);
  }
  
  // Formatting
  lbSheet.autoResizeColumns(1, numCols);
  if (lbData.length > 0) {
    lbSheet.getRange(2, 6, lbData.length, 1).setNumberFormat('yyyy-mm-dd');
  }
  
  // Add conditional formatting for tiers (only if there are players)
  lbSheet.setConditionalFormatRules([]);
  
  if (lbData.length > 0) {
    const ratingRange = lbSheet.getRange(2, 3, lbData.length, 1); // Column C (Rating)
    
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
  }
  
  // Alert completion
  SpreadsheetApp.getUi().alert(
    'Ratings Updated', 
    `Processed ${sortedPlayers.length} players across ${sortedMatchIDs.length} matches.`, 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏆 Refresh Leaderboard')
    .addItem('Calculate Ratings', 'calculateAllRatings')
    .addToUi();
}