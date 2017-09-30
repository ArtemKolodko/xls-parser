const FILE_NAME = 'book1.xlsx';
const RESULT_CELL_INDEX = 'D';
const TARGET_CELL_INDEX = 'AF';
const PROJECT_COLUMN = 3;
const PROGRAM_COLUMN = 'B';
const LIMIT = 500000;

let Excel = require('exceljs');

// read from a file 
let workbook = new Excel.Workbook();

let Program = function({indent=0, value, name, from, to}) {
    let self = this;
    this.programs = [];
    this.projects = [];

    this.indent = indent;
    this.value = value;
    this.name = name;

    this.from = from;
    this.to = to;

    this.addProgram = function(program) {
        self.programs.push(program)
    }

    this.addProject = function(project) {
        self.projects.push(project)
    }
}

function parseWorksheet(worksheet) {
    let programs = [];
	let lastProgram = new Program({indent: -1, name: 'entry point'});
	// Find parent node in project column
	let programCol = worksheet.getColumn(PROGRAM_COLUMN);
	let projectCol = worksheet.getColumn(PROJECT_COLUMN);
	programCol.eachCell(function(programCell, programRowNum) {

		if(programRowNum < 11 || programRowNum > LIMIT) {
			return false;
		}

		if(programCell.value === null) {

			let lastProgram = programs[programs.length - 1];

            projectCol.eachCell(function(projectCell, projectRowNum) {
            	if(projectRowNum < lastProgram.from+1) {
            		return false;
				}

				if(projectCell.value === null) {
            		return false;
				}

				if(projectRowNum > LIMIT) {
            		return false;
				}

                if(typeof lastProgram.to !== 'undefined') {
                    return false;
                }

				if(projectCell.value === 'Результат') {
                    lastProgram.to = projectRowNum-1;
                    return false;
				}

                let name = projectCell.value;
                let row = worksheet.getRow(projectRowNum);
                let indent = lastProgram.indent;
                let value = row.getCell(32).value;
                let newProject = new Program({
                    value,
                    name,
                    indent
                })

                //console.log('----Project', name, '\n');

                lastProgram.addProject(newProject);
			});

			return false;
		}

		let indent = programCell.alignment.indent || 0;

        // ищем сумму
        let row = worksheet.getRow(programRowNum);
        let value = row.getCell(32).value;
        let name = programCell.value;

        let newProgram = new Program({
            indent,
            name,
            value,
			from: programRowNum
        });

        //console.log('add program', name, programRowNum, '\n');

        programs.push(newProgram);
	})

	return programs;
}

function createWorkBook(programs) {
    let workbook = new Excel.Workbook();
    workbook.creator = 'Oscar the cat';
    workbook.lastModifiedBy = 'Oscar the cat';
    workbook.created = new Date();
    workbook.modified = new Date();

    let worksheet = workbook.addWorksheet('РЖД - наше все!', {
    	properties:{
    		tabColor:{argb:'FFC0000'}
    	}
    });
    worksheet.columns = [
        { header: 'Уровень', key: 'depth', width: 10 },
        { header: 'Программа', key: 'program', width: 80 },
        { header: 'Проект', key: 'project', width: 100 },
        { header: 'Плановые инвест. затраты (2017)', key: 'expenses', width: 50, outlineLevel: 1 }
    ];

    programs.forEach(function(p) {

    	//console.log('projects: ', typeof p.projects, p.projects.length);
        worksheet.addRow({depth: p.indent, program: p.name, project: null, expenses: p.value});

    	if(p.projects.length > 0) {
            p.projects.forEach(function(project) {
                worksheet.addRow({depth: project.indent, program: null, project: project.name, expenses: project.value});
            })
		}


	});

    return workbook;
}

function writeXLSX(workbook) {
	let exportFile = 'export.xlsx';
    workbook.xlsx.writeFile(exportFile)
        .then(function() {
            console.log('Export successfull! File name:',exportFile);
        });
}



workbook.xlsx.readFile(FILE_NAME)
    .then(function(data) {
        // use workbook 
        return parseWorksheet(data.getWorksheet(1));
        //console.log(programs)
    })
	.then(function(data) {
		return createWorkBook(data);
	})
	.then(function(workbook) {
        writeXLSX(workbook);
	});