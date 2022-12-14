const fs = require('fs');
const excel = require('excel4node');
const csv = require('fast-csv');
const readline = require('readline');

let config;
let input_file_names;

// finds category name for given racer result, filters out any prefaces before actual category name
function getCategoryName(racer_result) {
    let index = racer_result[7].search(/[a-zA-Z]/);

    return racer_result[7].substring(index);
}

// removes all categories from results that have no racers
function sanitizeCatNames(category_names, cat_results) {
    let to_be_removed = [];

    for (let cat_name of category_names) {
        if (!cat_results[cat_name]) {
            to_be_removed.push(category_names.indexOf(cat_name));
        }
    }

    for (let i = to_be_removed.length -1; i >= 0; i--) {
        category_names.splice(to_be_removed[i], 1);
    }
}

// uses first line of csv to get stage name and discards it
function getStageNames(race_times) {
    let csv_header = race_times[0];
    let stage_names = csv_header.slice(31, csv_header.indexOf('NumSplits'));

    stage_names = stage_names.map(stage_name => stage_name.substring(0, 2));

    for (let i = 0; i < stage_names.length; i++) {
        stage_names[i] = (i % 2 === 0) ? stage_names[i] + 'T' : stage_names[i] + 'P';
    }

    race_times.shift();

    return stage_names;
}

// covers cases for both DNF and N/C
function getPlace(result) {
    if (result[12] != '') {
        return result[12];
    } else if (result[11] == 'Y') {
        return 'N/C';
    } else if (result[13] == 'DNF') {
        return 'DNF';
    }
}

// main processing function
function processTimes(race_results, file_name) {
    console.log('Race results fetched, beginning processing...');

    const stage_names = getStageNames(race_results);
    const num_stages = stage_names.length / 2;
    const cat_results = {};

    for (let result of race_results) {
        let category_name = getCategoryName(result);

        let racer_json = {
            plate_num: result[0],
            place: getPlace(result),
            name: result[3],
            team_sponsor: result[5],
            overall_time: result[10],
            time_behind: result[25],
            stage_results: []
        }

        for (let i = 31; i < (num_stages * 2) + 31; i += 2) {
            racer_json.stage_results.push({
                stage_num: stage_names[i-31].substring(1, 2),
                stage_time: result[i].substring(3),
                stage_place: result[i+1]
            })
        }

        racer_json.stage_results.sort((a, b) => {
            if (a.stage_num < b.stage_num) {
                return -1;
            } else if (a.stage_num > b.stage_num) {
                return 1;
            } else {
                return 0;
            }
        });

        if (cat_results[category_name]) {
            cat_results[category_name].push(racer_json);
        } else {
            cat_results[category_name] = [racer_json];
        }
    }

    stage_names.sort((a, b) => {
        let a_num = parseInt(a.substring(1, 2));
        let b_num = parseInt(b.substring(1, 2)); 
        if (a_num < b_num) {
            return -1;
        } else if (a_num > b_num) {
            return 1;
        } else {
            return 0;
        }
    });

    generateExcel(cat_results, stage_names, file_name);
}

// generates excel from processed results
function generateExcel(cat_results, stage_names, file_name) {
    console.log('Result processing for ' + file_name + ' finished, generating output excel');

    const num_stages = stage_names.length / 2;
    const wb = new excel.Workbook();

    // configuring style of each section
    const title_style = wb.createStyle({
        font: {
            color: config.title_style.font_color,
            size: config.title_style.font_size,
            bold: config.title_style.is_bold,
            name: config.title_style.font_name
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: config.title_style.fill_color
        },
        border: {
            left: {
                style: config.title_style.border_thickness,
                color: '#000000'
            },
            right: {
                style: config.title_style.border_thickness,
                color: '#000000'
            },
            top: {
                style: config.title_style.border_thickness,
                color: '#000000'
            },
            bottom: {
                style: config.title_style.border_thickness,
                color: '#000000'
            }
        }
    });

    const header_style = wb.createStyle({
        font: {
            color: config.header_style.font_color,
            size: config.header_style.font_size,
            bold: config.header_style.is_bold,
            name: config.header_style.font_name
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: config.header_style.fill_color
        },
        border: {
            left: {
                style: config.header_style.border_thickness,
                color: '#000000'
            },
            right: {
                style: config.header_style.border_thickness,
                color: '#000000'
            },
            top: {
                style: config.header_style.border_thickness,
                color: '#000000'
            },
            bottom: {
                style: config.header_style.border_thickness,
                color: '#000000'
            }
        }
    });
    
    const body_style = wb.createStyle({
        font: {
            color: '#000000',
            size: config.body_style.font_size,
            name: config.body_style.font_name
        },
        border: {
            left: {
                style: config.body_style.border_thickness,
                color: '#000000'
            },
            right: {
                style: config.body_style.border_thickness,
                color: '#000000'
            },
            top: {
                style: config.body_style.border_thickness,
                color: '#000000'
            },
            bottom: {
                style: config.body_style.border_thickness,
                color: '#000000'
            }
        }
    });

    let ws = wb.addWorksheet('Results');

    // generate header
    ws.cell(1, 1)
        .string('**event title**')
        .style(title_style);
    
    for (let i = 2; i <= 7 + (num_stages * 2); i++) {
        ws.cell(1, i)
            .style(title_style);
    }

    // generate tables
    let currRow = 2;
    
    const category_names = [...config.categories];

    sanitizeCatNames(category_names, cat_results);

    for (let category_name of category_names) {
        // category name row
        ws.cell(currRow, 1)
            .string(category_name.toUpperCase())
            .style(title_style);

            for (let i = 2; i <= 7 + (num_stages * 2); i++) {
                ws.cell(currRow, i)
                    .style(title_style);
            }

        currRow++;

        // header row
        ws.cell(currRow, 1)
            .string('Place')
            .style(header_style);
        
        ws.cell(currRow, 2)
            .string('Points Earned')
            .style(header_style);

        ws.cell(currRow, 3)
            .string('Plate')
            .style(header_style);

        ws.cell(currRow, 4)
            .string('Name')
            .style(header_style);

        ws.cell(currRow, 5)
            .string('Category / Sponsor(s)')
            .style(header_style);

        ws.cell(currRow, 6)
            .string('Overall')
            .style(header_style);

        ws.cell(currRow, 7)
            .string('Behind')
            .style(header_style);

        for (let i = 8; i < 8 + (num_stages * 2); i += 2) {
            ws.cell(currRow, i)
                .string(stage_names[i - 8])
                .style(header_style);

            ws.cell(currRow, i + 1)
                .string(stage_names[i - 8 + 1])
                .style(header_style);
        }

        currRow++;
        
        // one row for each racer result per category
        for (let result of cat_results[category_name]) {
            ws.cell(currRow, 1)
                .string(result.place)
                .style(body_style);
            
            ws.cell(currRow, 2)
                .number(config.points_table[result.place])
                .style(body_style);

            ws.cell(currRow, 3)
                .string(result.plate_num)
                .style(body_style);

            ws.cell(currRow, 4)
                .string(result.name)
                .style(body_style);

            ws.cell(currRow, 5)
                .string(result.team_sponsor)
                .style(body_style);

            ws.cell(currRow, 6)
                .string(result.overall_time.substring(1))
                .style(body_style);
            
            ws.cell(currRow, 7)
                .string(result.time_behind)
                .style(body_style);

            for (let i = 8; i < 8 + (num_stages * 2); i += 2) {
                ws.cell(currRow, i)
                    .string(result.stage_results[(i - 8) / 2].stage_time)
                    .style(body_style);
                ws.cell(currRow, i + 1)
                    .string(result.stage_results[(i - 8) / 2].stage_place)
                    .style(body_style);
            }

            currRow++;
        }

        currRow++;
    }

    wb.write(file_name.substring(0, file_name.indexOf('.csv')) + '.xlsx');

    console.log(file_name + ' done!');
}

// initial setting of input_file_names
function setup() {
    config = JSON.parse(fs.readFileSync('./config.json', 'utf-8'));

    if (config.specify_input_file) {
        input_file_names = config.input_file_names;
    } else {
        input_file_names = fs.readdirSync('./').filter(file => file.includes('.csv'));
    }
}

// handles csv import and kicks off processing for each
function main() {
    const readline_interface = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    const race_times = {};

    console.log('Starting processor, fetching input files...');
    console.log('Found file(s): ' + input_file_names);

    for (let file_name of input_file_names) {
        if (!fs.existsSync(file_name)) {
            console.error('ERROR - file with name \"' + file_name + '\" does not exist, moving to next file');
        } else {
            race_times[file_name] = [];
            fs.createReadStream(file_name).pipe(csv.parse({ headers: false }))
                .on('error', error => console.error(error))
                .on('data', row => race_times[file_name].push(row))
                .on('end', () => processTimes(race_times[file_name], file_name));
        }
    }
}                                                                                                                  

setup();
main();