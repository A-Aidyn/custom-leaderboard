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
    
    // Performance sharpening constants
    const PERF_GAMMA = 2.5;  // exponent to widen spread between high/low performers
    const PERF_MIN = 0.70;   // clamp floor for perfIndex before sharpening
    const PERF_MAX = 1.90;   // clamp ceiling for perfIndex before sharpening
    
    // Pass 1: compute per-player stats for each team
    const computeTeamStats = (team, myTeamPerf, myTeamAvgRating, oppTeamAvgRating) => {
      const expected = 1 / (1 + Math.pow(10, (oppTeamAvgRating - myTeamAvgRating) / 400));  // expected range: (0, 1)
      const baseChange = myTeamPerf - expected;  // baseChange range: (-1, 1)
      const isGain = baseChange > 0;
      
      return team.map(p => {
        const state = players[p.player];
        
        // Individual performance metrics
        const myKDA = (p.kills + p.assists * 0.5) / Math.max(p.deaths, 1);  // myKDA range: [0, ...]
        const acsRatio = lobbyACS > 0 ? p.acs / lobbyACS : 1;  // acsRatio range: [0, 10]
        const kdaRatio = lobbyKDA > 0 ? Math.min(myKDA / lobbyKDA, 2.5) : 1;  // kdaRatio range: [0, 2.5]
        const perfIndex = 0.6 * acsRatio + 0.4 * kdaRatio;  // perfIndex range: [0, 7]
        
        // Sharpen: clamp then raise to PERF_GAMMA to widen spread
        const perfClamped = Math.max(PERF_MIN, Math.min(PERF_MAX, perfIndex));
        const rawPerf = Math.pow(perfClamped, PERF_GAMMA);  // rawPerf range: [0.70^2.5, 1.90^2.5] ≈ [0.41, 4.97]
        
        // K-factor with uncertainty
        const baseK = 32 * Math.max(1 - state.matches / 30, 0.5);  // baseK range: [16, 32]
        var u01 = Math.max(0, Math.min(1, (state.uncertainty - U_MIN) / (U_MAX - U_MIN)));  // u01 range: [0, 1]
        var kMult = 1 + u01 * (K_MULT_MAX - 1);  // kMult range: [1, 2]
        var kFactor = baseK * kMult;  // kFactor range: [16, 64]
        
        return { p: p, state: state, perfIndex: perfIndex, rawPerf: rawPerf, kFactor: kFactor, baseChange: baseChange, isGain: isGain };
      });
    };
    
    // Pass 2: normalize within team and apply rating changes
    const applyTeamChanges = (teamStats) => {
      // K-weighted mean of rawPerf (for wins) or 1/rawPerf (for losses)
      var sumK = teamStats.reduce((s, t) => s + t.kFactor, 0);
      
      if (teamStats[0].isGain) {
        var meanRaw = teamStats.reduce((s, t) => s + t.kFactor * t.rawPerf, 0) / sumK;
        teamStats.forEach(t => { t.perfMod = t.rawPerf / meanRaw; });  // normalized so K-weighted average = 1
      } else {
        var meanInv = teamStats.reduce((s, t) => s + t.kFactor * (1 / t.rawPerf), 0) / sumK;
        teamStats.forEach(t => { t.perfMod = (1 / t.rawPerf) / meanInv; });  // high performers lose less
      }
      
      teamStats.forEach(t => {
        var ratingChange = t.kFactor * t.baseChange * t.perfMod;
        
        t.state.rating += ratingChange;
        t.state.matches++;
        t.state.uncertainty = Math.max(U_MIN, t.state.uncertainty * DECAY);
        t.state.lastDate = matchDate;
        
        rowRatingChanges[t.p.rowIndex] = Math.round(ratingChange);
        
        Logger.log(`${t.p.player}: ${(t.baseChange + (t.isGain ? 0.5 : 0.5)).toFixed(3)} team, ` +
                    `PI=${t.perfIndex.toFixed(3)}, perfMod=${t.perfMod.toFixed(3)}, Δ=${ratingChange.toFixed(2)}, ` +
                    `new=${t.state.rating.toFixed(2)}`);
      });
    };
    
    // Process both teams
    applyTeamChanges(computeTeamStats(teamA, teamAPerf, teamAAvgRating, teamBAvgRating));
    applyTeamChanges(computeTeamStats(teamB, teamBPerf, teamBAvgRating, teamAAvgRating));
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