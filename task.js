const request = require("request");
const cheerio = require("cheerio");
const xl = require('excel4node');

request("https://www.espncricinfo.com/series/ipl-2020-21-1210595/match-results", cb);

function cb(err, response, html) {
    getScorecard(html);
}

function getScorecard(html) {
    let $ = cheerio.load(html);
    let ScoreCardElement = $(".btn-sm.btn-outline-dark.match-cta");
    var matches = 0;
    for (let i = 2; i < ScoreCardElement.length; i = i + 4) {              //ScoreCardElement.length
        let href = $(ScoreCardElement[i]).attr('href');
        let FullLink = "https://www.espncricinfo.com" + href;
        matches++;
        statsHTML(FullLink, matches);
    }
}

function statsHTML(FullLink, matches) {
    request(FullLink, cb);
    function cb(err, response, html) {
        Stats(html, matches);
    }
}

function Stats(html, matches) {
    let $ = cheerio.load(html);
    //  For batsMan 

    let ScoreCard = $(".batsman.table");
    var wb = new xl.Workbook();
    let result = $(".match-info.match-info-MATCH.match-info-MATCH-half-width>.status-text").text();

    for (var s = 0; s < ScoreCard.length; s++) {
        var ws = wb.addWorksheet('BatsMan Inning - ' + s);
        let NamesA = [];
        let RunsA = [];
        let BallA = [];
        let FourA = [];
        let SixA = [];
        let StrikeRA = [];

        let BatsmanTable = $(ScoreCard[s]).find('tr');

        let total = $(BatsmanTable[BatsmanTable.length - 3]).text();
        let NotBat = $(BatsmanTable[BatsmanTable.length - 2]).text();

        for (var i = 0; i < BatsmanTable.length - 4; i++) {
            if ($(BatsmanTable[i]).text() != "") {
                let Rows = $(BatsmanTable[i]).find('td');
                let Names = $(BatsmanTable[i]).find('a.small').text();
                let Runs = $(BatsmanTable[i]).find('.font-weight-bold').text();
                let Balls = ($(Rows[3]).text());
                let Fours = ($(Rows[5]).text());
                let Sixes = ($(Rows[6]).text());
                let StrikeRate = ($(Rows[7]).text());

                NamesA.push(Names);
                RunsA.push(Runs);
                BallA.push(Balls);
                FourA.push(Fours);
                SixA.push(Sixes);
                StrikeRA.push(StrikeRate);

                ws.cell(1, 1).string(result);
                ws.cell(2, 1).string(total);
                ws.cell(3, 1).string(NotBat);

                ws.cell(4, 1).string("Name");
                ws.cell(4, 2).string("Runs");
                ws.cell(4, 3).string("Balls Played");
                ws.cell(4, 4).string("Fours");
                ws.cell(4, 5).string("Sixs");
                ws.cell(4, 6).string("Strike Rate")
            }
        }
        for (var j = 1; j < (BatsmanTable.length - 3) / 2; j++) {
            let Name = NamesA[j];
            let Run = RunsA[j];
            let Ball = BallA[j];
            let Four = FourA[j];
            let Six = SixA[j];
            let StrikeA = StrikeRA[j];

            let k = j + 4;
            ws.cell(k, 1).string(Name);
            ws.cell(k, 2).string(Run);
            ws.cell(k, 3).string(Ball);
            ws.cell(k, 4).string(Four);
            ws.cell(k, 5).string(Six);
            ws.cell(k, 6).string(StrikeA);
        }
    }
    // For Bowlers
    let BowlerCard = $(".bowler.table");
    for (var b = 0; b < BowlerCard.length; b++) {
        var ws = wb.addWorksheet('Bowler Inning - ' + b);
        let emptyRowCount = 0;

        let NamesB = [];
        let OversB = [];
        let MaidenB = [];
        let RunsB = [];
        let WicketB = [];
        let EcoB = [];
        let DotB = [];
        let FourB = [];
        let SixB = [];
        let WideB = [];
        let NoB = [];

        let BowlerTable = $(BowlerCard[b]).find('tr');
        for (var i = 0; i < BowlerTable.length; i++) {
            if ($(BowlerTable[i]).text() != "") {
                let Rows = $(BowlerTable[i]).find('td');
                let Names = $(BowlerTable[i]).find('a.small').text();
                let Overs = ($(Rows[1]).text());
                let Maiden = ($(Rows[2]).text());
                let Runs = ($(Rows[3]).text());
                let Wicket = ($(Rows[4]).text());
                let Eco = ($(Rows[5]).text());
                let Dot = ($(Rows[6]).text());
                let Four = ($(Rows[7]).text());
                let Six = ($(Rows[8]).text());
                let Wide = ($(Rows[9]).text());
                let No = ($(Rows[10]).text());

                NamesB.push(Names);
                OversB.push(Overs);
                MaidenB.push(Maiden);
                RunsB.push(Runs);
                WicketB.push(Wicket);
                EcoB.push(Eco);
                DotB.push(Dot);
                FourB.push(Four);
                SixB.push(Six);
                WideB.push(Wide);
                NoB.push(No);

                ws.cell(1, 1).string(result);
                ws.cell(2, 1).string("Full Name");
                ws.cell(2, 2).string("Overs");
                ws.cell(2, 3).string("Maiden Overs");
                ws.cell(2, 4).string("Runs");
                ws.cell(2, 5).string("Wickets Taken");
                ws.cell(2, 6).string("ECON");
                ws.cell(2, 7).string("Dot Ball");
                ws.cell(2, 8).string("Fours");
                ws.cell(2, 9).string("Sixes");
                ws.cell(2, 10).string("Wide Ball");
                ws.cell(2, 11).string("No Ball");
            }
            else {
                emptyRowCount++;
            }
        }
        for (var j = 1; j < BowlerTable.length - emptyRowCount; j++) {
            let k = j + 2;
            let Name = NamesB[j];
            let Overs = OversB[j];
            let Maiden = MaidenB[j];
            let Runs = RunsB[j];
            let Wicket = WicketB[j];
            let ECON = EcoB[j];
            let Dot = DotB[j];
            let Four = FourB[j];
            let Six = SixB[j];
            let Wide = WideB[j];
            let No = NoB[j];

            ws.cell(k, 1).string(Name);
            ws.cell(k, 2).string(Overs);
            ws.cell(k, 3).string(Maiden);
            ws.cell(k, 4).string(Runs);
            ws.cell(k, 5).string(Wicket);
            ws.cell(k, 6).string(ECON);
            ws.cell(k, 7).string(Dot);
            ws.cell(k, 8).string(Four);
            ws.cell(k, 9).string(Six);
            ws.cell(k, 10).string(Wide);
            ws.cell(k, 11).string(No);
        }
    }
    wb.write('Match Number - ' + matches + '.xlsx');
}


// npm i request
// npm i cheerio
// npm i excel4node
// node task.js