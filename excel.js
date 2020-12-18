const Excel = require('exceljs');
const XlsxPopulate = require('xlsx-populate');
const Util = require('./utils');


const AssetType = {
    fixed: 1,  //основні засоби
    intangible: 2  //нематеріальні активи
}

const ActType = {
    WEAR: 'wear',
    WITHOUTWEAR: 'withoutwear',
    BABAK: 'babak'
}

function init() {
    let document = null;

    Repo.getAssetsFixed().then((assets) => {
        document = new Document(assets, ActType.WEAR);
        return document.create();
    })
        .then((assets) => {
            Repo.getSubjects().then((subjects) => {
                assets.forEach(asset => {
                    document.createTitle(ActType.WEAR, asset, subjects);
                })
            });
        });

    Repo.getAssetsIntangible().then((assets) => {
        document = new Document(assets, ActType.BABAK);
        return document.create();
    })
        .then((assets) => {
            Repo.getSubjects().then((subjects) => {
                assets.forEach(asset => {
                    document.createTitle(ActType.BABAK, asset, subjects);
                })
            });
        });
}

//---------------ASSET-----------------------------------------//
let AssetFactory = {
    create: (type, name) => {
        switch (type) {
            case AssetType.fixed: {
                return new AssetFixedData(name);
            }
            case AssetType.intangible: {
                return new AssetIntangibleData(name);
            }
            default:
                return null;
        }
    }
}

let AssetFixedData = function (name) {
    this.name = name;
    this.type = AssetType.fixed;
    let _data = [];

    this.addAsset = (value) => {
        _data.push(value);
    }

    this.getData = () => {
        return _data;
    }

    this.getTotal = () => {
        return {
            count: _data.length,
            amount: _data.length,
            startcost: _data.reduce((sum, current) => { return sum + current.startcost }, 0),
            endcost: _data.reduce((sum, current) => { return sum + current.endcost }, 0)
        };
    }
}

let AssetIntangibleData = function (name) {
    this.name = name;
    this.type = AssetType.intangible;
    let _data = [];

    this.addAsset = (value) => {
        _data.push(value);
    }

    this.getData = () => {
        return _data;
    }

    this.getTotal = () => {
        return {
            count: _data.length,
            amount: _data.reduce((c, current) => { return c + current.amount }, 0),
            cost: _data.reduce((sum, current) => { return sum + current.cost }, 0),
            sum: _data.reduce((sum, current) => { return sum + current.sum }, 0)
        };
    }
}

let AssetFixedModel = function ({ code, name, date, startcost, wear, endcost, bill }) {
    this.code = code;
    this.name = name;
    this.date = date;
    this.startcost = startcost;
    this.wear = wear;
    this.endcost = endcost;
    this.bill = bill;
}

let AssetIntangibleModel = function ({ bill, name, unit, amount, cost, sum, fund }) {
    this.bill = bill;
    this.name = name;
    this.unit = unit;
    this.amount = amount;
    this.cost = cost;
    this.sum = sum;
    this.fund = fund;
}
//---------------END ASSET-----------------------------------------//

//---------------SUBJECT-----------------------------------------//

let Subject = function (code, name) {
    this.code = code;
    this.name = name;
}
//---------------END SUBJECT-----------------------------------------//

//---------------AREA-----------------------------------------//
let PageModel = function (page, hb, he, cb, ce, fb, fe) {
    this.page = page;
    this.headerBegin = hb;
    this.headerEnd = he;
    this.contentBegin = cb;
    this.contentEnd = ce;
    this.footerBegin = fb;
    this.footerEnd = fe;
}

let Page = function (actType = ActType.WEAR) {
    let _page = 1;
    let _headerHeight = 100;
    let _contentHeight = 600; //let _contentHeight = 540;
    let _footerHeight = 165;
    let _rowHeight = 15;
    let _area;

    function constructor() {
        switch (actType) {
            case ActType.WEAR: {
                _area = new PageModel(_page, 0, _headerHeight, _headerHeight + _rowHeight, 715, 730, 880);
                break;
            }
            case ActType.WITHOUTWEAR: {
                _area = new PageModel(_page, 0, _headerHeight, _headerHeight + _rowHeight, 715, 730, 880);
                break;
            }
            case ActType.BABAK: {
                _headerHeight = 60;
                _contentHeight = 640;
                _area = new PageModel(_page, 0, _headerHeight, _headerHeight + _rowHeight, 715, 730, 880);
                break;
            }
            default: {
                console.log('Act type not found!');
            }
        }

    }
    constructor();

    return {
        next: () => {
            _page = _page + 1;
            if (_page > 1) {
                let hb = _area.footerEnd + _rowHeight;
                let he = hb + _headerHeight;
                let cb = he + _rowHeight;
                let ce = cb + _contentHeight;
                let fb = ce + _rowHeight;
                let fe = (fb + _footerHeight) - _rowHeight;
                _area = new PageModel(_page, hb, he, cb, ce, fb, fe);
            }
        },
        getPage: () => {
            return _page;
        },
        getArea: () => {
            return _area;
        },
        getHeaderHeight: () => {
            return _headerHeight;
        },
        getFooterHeight: () => {
            return _footerHeight;
        },
        getRowHeight: () => {
            return _rowHeight;
        }
    }
}
//---------------END AREA-----------------------------------------//

let Repo = (function () {
    let _assets = [];

    return {
        getSubjects: () => {
            return XlsxPopulate.fromFileAsync('./template/subjects.xlsx')
                .then(function (workbook) {
                    let sheet = workbook.sheet('Теми');
                    let subjects = sheet.usedRange().value();
                    let data = [];
                    subjects.map(item => {
                        data.push(new Subject(item[0], item[1]));
                    });
                    _subjects = data;
                    return data;
                });
        },
        getAssetsFixed: async () => {
            let dataWorkbook = new Excel.Workbook();
            await dataWorkbook.xlsx.readFile('./template/fixed.xlsx');

            dataWorkbook.eachSheet((sheet, sheetId) => {
                let asset = AssetFactory.create(AssetType.fixed, sheet.name);
                sheet.eachRow((row, rowNumber) => {
                    if (row.hidden || row.number == 1) return;

                    let newAsset = new AssetFixedModel(
                        {
                            code: typeof (row.values[1]) === 'number' ? row.values[1] : row.values[1].replace(/\s+/g, ''),
                            // code: row.values[1],
                            name: row.values[2],
                            date: row.values[5],
                            startcost: row.values[6],
                            wear: row.values[7],
                            endcost: row.values[8],
                            bill: row.values[9]
                        }
                    );
                    asset.addAsset(newAsset);
                })
                _assets.push(asset);
            });
            return _assets;
        },
        getAssetsIntangible: async () => {
            let dataWorkbook = new Excel.Workbook();
            try {
                await dataWorkbook.xlsx.readFile('./template/intangible.xlsx');

                dataWorkbook.eachSheet((sheet, sheetId) => {
                    let asset = AssetFactory.create(AssetType.intangible, sheet.name);
                    sheet.eachRow((row, rowNumber) => {
                        if (row.hidden || row.number == 1) return;

                        let newAsset = new AssetIntangibleModel(
                            {
                                bill: row.values[1],
                                name: row.values[2],
                                unit: row.values[3],
                                amount: row.values[4],
                                cost: row.values[5],
                                sum: row.values[6],
                                fund: row.values[7]
                            }
                        );
                        asset.addAsset(newAsset);
                    })
                    _assets.push(asset);
                });
                return _assets;
            } catch (error) {
                console.log(`Error: Find out next error: ${error}`);
            }

        },
    }
})();

let Document = function (assets, actType) {
    let page = null;
    let newWorkbook = new Excel.Workbook();
    let sheet = newWorkbook.addWorksheet('Звіт', {
        pageSetup: {
            paperSize: 9,
            orientation: 'landscape',
            margins: {
                left: 0.4,
                top: 0.4,
                right: 0.2,
                bottom: 0.2,
                footer: 0,
                header: 0
            },
            scale: 60
        }
    });

    let height = 0;
    let total = null;

    insertHeader = () => {
        switch (actType) {
            case ActType.WEAR: {
                let lastRow = sheet.lastRow != undefined ? sheet.lastRow.number + 1 : 1;
                firstRow = lastRow;
                secondRow = lastRow + 1;
                sheet.insertRow(firstRow,
                    [
                        '№ з/п',
                        `Найменування, стисла характеристика та призначення об'єкта`,
                        `Рік випуску (будівництва) чи дата придбання (введення в експлуатацію) та виготовлювач`,
                        `Номер`,
                        ``,
                        ``,
                        `Одиниця виміру`,
                        `Фактична наявність`,
                        ``,
                        `Відмітка про вибуття`,
                        `За даними бухгалтерського обліку`,
                        '',
                        '',
                        '',
                        '',
                        `Інші відомості`
                    ]);
                sheet.insertRow(secondRow,
                    [
                        '',
                        '',
                        '',
                        'інвентарний/номенклатурний',
                        `заводський`,
                        `паспорта`,
                        '',
                        'кількість',
                        'первісна (переоцінена) вартість',
                        '',
                        'кількість',
                        'первісна (переоцінена) вартість',
                        'сума зносу (накопиченої амортизації)',
                        'балансова  вартість',
                        'строк корисного використання',
                        ''
                    ]);
                sheet.mergeCells(`D${firstRow}:F${firstRow}`);
                sheet.mergeCells(`H${firstRow}:I${firstRow}`);
                sheet.mergeCells(`K${firstRow}:O${firstRow}`);
                sheet.mergeCells(`A${firstRow}:A${secondRow}`);
                sheet.mergeCells(`B${firstRow}:B${secondRow}`);
                sheet.mergeCells(`C${firstRow}:C${secondRow}`);
                sheet.mergeCells(`G${firstRow}:G${secondRow}`);
                sheet.mergeCells(`J${firstRow}:J${secondRow}`);
                sheet.mergeCells(`P${firstRow}:P${secondRow}`);

                sheet.getColumn(1).width = 10;
                sheet.getColumn(2).width = 50;
                sheet.getColumn(3).width = 18;
                sheet.getColumn(4).width = 20;
                sheet.getColumn(5).width = 11;
                sheet.getColumn(6).width = 11;
                sheet.getColumn(7).width = 10;
                sheet.getColumn(8).width = 10;
                sheet.getColumn(9).width = 20;
                sheet.getColumn(10).width = 10;
                sheet.getColumn(11).width = 10;
                sheet.getColumn(12).width = 20;
                sheet.getColumn(13).width = 20;
                sheet.getColumn(14).width = 20;
                sheet.getColumn(15).width = 15;
                sheet.getColumn(16).width = 10;

                sheet.getRow(firstRow).font = { name: 'Times New Roman', size: 12, family: 4 };
                sheet.getRow(secondRow).font = { name: 'Times New Roman', size: 12, family: 4 };
                sheet.getCell(`A${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`B${firstRow}`).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                sheet.getCell(`C${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`D${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`G${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`I${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`J${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`K${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`L${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`M${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`N${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`O${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`P${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

                sheet.getCell(`A${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`B${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`C${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`D${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`G${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`I${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`J${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`K${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`L${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`M${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`N${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`O${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`P${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

                sheet.getCell(`A${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`B${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`C${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`D${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`E${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`F${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`G${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`H${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`I${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`J${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`K${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`L${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`M${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`N${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`O${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`P${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };

                sheet.getCell(`A${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`B${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`C${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`D${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`E${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`F${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`G${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`H${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`I${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`J${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`K${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`L${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`M${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`N${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`O${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`P${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getRow(secondRow).height = 100;
                break;
            }
            case ActType.WITHOUTWEAR: {
                let lastRow = sheet.lastRow != undefined ? sheet.lastRow.number + 1 : 1;
                firstRow = lastRow;
                secondRow = lastRow + 1;
                sheet.insertRow(firstRow,
                    [
                        '№ з/п',
                        `Найменування, стисла характеристика та призначення об'єкта`,
                        `Рік випуску (будівництва) чи дата придбання (введення в експлуатацію) та виготовлювач`,
                        `Номер`,
                        ``,
                        ``,
                        `Одиниця виміру`,
                        `Фактична наявність`,
                        ``,
                        `Відмітка про вибуття`,
                        `За даними бухгалтерського обліку`,
                        ``,
                        `Інші відомості`
                    ]);
                sheet.insertRow(secondRow,
                    [
                        '',
                        '',
                        '',
                        'інвентарний/номенклатурний',
                        `заводський`,
                        `паспорта`,
                        '',
                        'кількість',
                        'первісна (переоцінена) вартість',
                        '',
                        'кількість',
                        'первісна (переоцінена) вартість',
                        '',
                        '',
                    ]);
                sheet.mergeCells(`D${firstRow}:F${firstRow}`);
                sheet.mergeCells(`H${firstRow}:I${firstRow}`);
                sheet.mergeCells(`K${firstRow}:L${firstRow}`);
                sheet.mergeCells(`A${firstRow}:A${secondRow}`);
                sheet.mergeCells(`B${firstRow}:B${secondRow}`);
                sheet.mergeCells(`C${firstRow}:C${secondRow}`);
                sheet.mergeCells(`G${firstRow}:G${secondRow}`);
                sheet.mergeCells(`J${firstRow}:J${secondRow}`);
                sheet.mergeCells(`M${firstRow}:M${secondRow}`);

                sheet.getColumn(1).width = 10;
                sheet.getColumn(2).width = 50;
                sheet.getColumn(3).width = 18;
                sheet.getColumn(4).width = 20;
                sheet.getColumn(5).width = 15;
                sheet.getColumn(6).width = 15;
                sheet.getColumn(7).width = 10;
                sheet.getColumn(8).width = 15;
                sheet.getColumn(9).width = 20;
                sheet.getColumn(10).width = 10;
                sheet.getColumn(11).width = 15;
                sheet.getColumn(12).width = 20;
                sheet.getColumn(13).width = 10;

                sheet.getRow(firstRow).font = { name: 'Times New Roman', size: 12, family: 4 };
                sheet.getRow(secondRow).font = { name: 'Times New Roman', size: 12, family: 4 };
                sheet.getCell(`A${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`B${firstRow}`).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                sheet.getCell(`C${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`D${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`G${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`I${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`J${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`K${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`L${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`M${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

                sheet.getCell(`A${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`B${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`C${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`D${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`G${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`I${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`J${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`K${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`L${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`M${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

                sheet.getCell(`A${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`B${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`C${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`D${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`E${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`F${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`G${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`H${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`I${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`J${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`K${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`L${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`M${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };

                sheet.getCell(`A${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`B${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`C${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`D${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`E${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`F${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`G${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`H${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`I${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`J${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`K${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`L${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`M${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getRow(secondRow).height = 90;
                break;
            }
            case ActType.BABAK: {
                let lastRow = sheet.lastRow != undefined ? sheet.lastRow.number + 1 : 1;
                firstRow = lastRow;
                secondRow = lastRow + 1;
                sheet.insertRow(firstRow,
                    [
                        '№ з/п',
                        'Рахунок, суб-рахунок',
                        'Матеріальні цінності',
                        '',
                        'Одиниця виміру',
                        'Фактична наявність',
                        '',
                        '',
                        `За даними бухгалтерського обліку`,
                        '',
                        '',
                        'Інші відомості'
                    ]);
                sheet.insertRow(secondRow,
                    [
                        '',
                        '',
                        'найменування, вид, сорт, група',
                        'номенклатурний номер (за наявності)',
                        '',
                        'кількість',
                        'вартість',
                        'сума',
                        'кількість',
                        'вартість',
                        'сума',
                        ''
                    ]);
                sheet.mergeCells(`C${firstRow}:D${firstRow}`);
                sheet.mergeCells(`F${firstRow}:H${firstRow}`);
                sheet.mergeCells(`I${firstRow}:K${firstRow}`);

                sheet.mergeCells(`A${firstRow}:A${secondRow}`);
                sheet.mergeCells(`B${firstRow}:B${secondRow}`);
                sheet.mergeCells(`E${firstRow}:E${secondRow}`);
                sheet.mergeCells(`L${firstRow}:L${secondRow}`);

                sheet.getColumn(1).width = 10;
                sheet.getColumn(2).width = 10;
                sheet.getColumn(3).width = 50;
                sheet.getColumn(4).width = 18;
                sheet.getColumn(5).width = 10;
                sheet.getColumn(6).width = 15;
                sheet.getColumn(7).width = 20;
                sheet.getColumn(8).width = 20;
                sheet.getColumn(9).width = 15;
                sheet.getColumn(10).width = 20;
                sheet.getColumn(11).width = 20;
                sheet.getColumn(12).width = 15;

                sheet.getRow(firstRow).font = { name: 'Times New Roman', size: 12, family: 4 };
                sheet.getRow(secondRow).font = { name: 'Times New Roman', size: 12, family: 4 };
                sheet.getCell(`A${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`B${firstRow}`).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                sheet.getCell(`C${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`D${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`G${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`I${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`J${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`K${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`L${firstRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

                sheet.getCell(`A${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`B${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`C${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`D${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`G${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`I${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`J${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`K${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`L${secondRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

                sheet.getCell(`A${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`B${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`C${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`D${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`E${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`F${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`G${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`H${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`I${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`J${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`K${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`L${firstRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };

                sheet.getCell(`A${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`B${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`C${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`D${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`E${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`F${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`G${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`H${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`I${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`J${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`K${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`L${secondRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getRow(secondRow).height = 60;
                break;
            }
            default: {
                console.log(`Unknown act type. From function 'insertHeader'`);
                break;
            }
        }
    }

    insertRow = (value) => {
        switch (actType) {
            case ActType.WITHOUTWEAR: {
                let lastRow = sheet.lastRow != undefined ? sheet.lastRow.number + 1 : 1;
                sheet.insertRow(lastRow, value, 'o');

                sheet.getRow(lastRow).font = { name: 'Times New Roman', size: 12, family: 4 };
                sheet.getCell(`A${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`B${lastRow}`).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                sheet.getCell(`C${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`D${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`G${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`I${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`J${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`K${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`L${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`M${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

                sheet.getCell(`A${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`B${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`C${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`D${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`E${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`F${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`G${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`H${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`I${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`J${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`K${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`L${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`M${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                break;
            }
            case ActType.WEAR: {
                let lastRow = sheet.lastRow != undefined ? sheet.lastRow.number + 1 : 1;
                sheet.insertRow(lastRow, value, 'o');

                sheet.getColumn(9).numFmt = '#,##0.00';
                sheet.getColumn(12).numFmt = '#,##0.00';
                sheet.getColumn(13).numFmt = '#,##0.00';
                sheet.getColumn(14).numFmt = '#,##0.00';

                sheet.getRow(lastRow).font = { name: 'Times New Roman', size: 12, family: 4 };
                sheet.getCell(`A${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`B${lastRow}`).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                sheet.getCell(`C${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`D${lastRow}`).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                sheet.getCell(`E${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`G${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`I${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`J${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`K${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`L${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`M${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`N${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`O${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`P${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

                sheet.getCell(`A${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`B${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`C${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`D${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`E${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`F${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`G${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`H${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`I${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`J${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`K${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`L${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`M${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`N${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`O${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`P${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                break;
            }
            case ActType.BABAK: {
                let lastRow = sheet.lastRow != undefined ? sheet.lastRow.number + 1 : 1;
                sheet.insertRow(lastRow, value);

                sheet.getColumn(7).numFmt = '#,##0.00';
                sheet.getColumn(8).numFmt = '#,##0.00';
                sheet.getColumn(10).numFmt = '#,##0.00';
                sheet.getColumn(11).numFmt = '#,##0.00';

                sheet.getRow(lastRow).font = { name: 'Times New Roman', size: 12, family: 4 };
                sheet.getCell(`A${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`B${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`C${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`D${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`G${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`I${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`J${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`K${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`L${lastRow}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

                sheet.getCell(`A${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`B${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`C${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`D${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`E${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`F${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`G${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`H${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`I${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`J${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`K${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                sheet.getCell(`L${lastRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
                break;
            }
            default: {
                console.log(`Unknown act type. From function 'insertRow'`);
                break;
            }
        }
    }

    insertFooter = (value) => {
        switch (actType) {
            case ActType.WEAR: {
                let row = sheet.lastRow != undefined ? sheet.lastRow.number + 1 : 1;
                sheet.insertRow(row,
                    ['Разом:',
                        '',
                        'x',
                        'x',
                        'x',
                        'x',
                        'x',
                        value.amount,
                        value.startcost,
                        'x',
                        value.amount,
                        value.startcost,
                        value.wear,
                        value.endcost,
                        'x',
                        'x',
                    ]);
                sheet.mergeCells(`A${row}:B${row}`);
                sheet.getRow(row).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`C${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`D${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`E${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`F${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`G${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`H${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`I${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`J${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`K${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`L${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`M${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`N${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`O${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`P${row}`).alignment = { vertical: 'middle', horizontal: 'center' };

                sheet.getCell(`A${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`C${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`D${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`E${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`F${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`G${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`H${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`I${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`J${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`K${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`L${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`M${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`N${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`O${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`P${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

                sheet.getCell(`I${row}`).numFmt = '#,##0.00';
                sheet.getCell(`L${row}`).numFmt = '#,##0.00';
                sheet.getCell(`M${row}`).numFmt = '#,##0.00';
                sheet.getCell(`N${row}`).numFmt = '#,##0.00';

                sheet.insertRow(row + 1, []);
                sheet.insertRow(row + 2, [
                    'Разом по сторінці:',
                    '',
                    'кількість порядкових номерів',
                    '',
                    `${Util.numberToDigit(value.amount)} ${Util.numberToString(value.amount, false)}`,
                    '', '', '', '', '', '', '', ''
                ]);
                sheet.mergeCells(`A${row + 2}:B${row + 2}`);
                sheet.mergeCells(`C${row + 2}:D${row + 2}`);
                sheet.mergeCells(`E${row + 2}:M${row + 2}`);
                sheet.getRow(row + 2).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`E${row + 2}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${row + 2}`).border = { bottom: { style: 'thin' } };
                sheet.insertRow(row + 3, [
                    '', '', '', '', '(прописом)', '', '', '', '', '', '', '', ''
                ]);
                sheet.getRow(row + 3).font = { name: 'Times New Roman', size: 12, family: 4, italic: true };
                sheet.mergeCells(`E${row + 3}:M${row + 3}`);
                sheet.getCell(`E${row + 3}`).alignment = { vertical: 'top', horizontal: 'center' };

                //---------------------------------------//
                sheet.insertRow(row + 4, [
                    '',
                    '',
                    'загальна кількість одиниць (фактично)',
                    '', '',
                    `${Util.numberToDigit(value.amount)} ${Util.numberToString(value.amount, false)}`,
                    '', '', '', '', '', '', ''
                ]);
                sheet.mergeCells(`A${row + 4}:B${row + 4}`);
                sheet.mergeCells(`C${row + 4}:E${row + 4}`);
                sheet.mergeCells(`F${row + 4}:M${row + 4}`);
                sheet.getRow(row + 4).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`F${row + 4}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${row + 4}`).border = { bottom: { style: 'thin' } };
                sheet.insertRow(row + 5, [
                    '', '', '', '', '', '(прописом)', '', '', '', '', '', '', ''
                ]);
                sheet.getRow(row + 5).font = { name: 'Times New Roman', size: 12, family: 4, italic: true };
                sheet.mergeCells(`F${row + 5}:M${row + 5}`);
                sheet.getCell(`F${row + 5}`).alignment = { vertical: 'top', horizontal: 'center' };

                //---------------------------------------//
                sheet.insertRow(row + 6, [
                    '',
                    '',
                    'вартість фактична',
                    '',
                    `${Util.numberToDigit(value.startcost.toFixed(2))} ${Util.moneyToString(value.startcost)}`,
                    '', '', '', '', '', '', '', ''
                ]);
                sheet.mergeCells(`A${row + 6}:B${row + 6}`);
                sheet.mergeCells(`C${row + 6}:D${row + 6}`);
                sheet.mergeCells(`E${row + 6}:M${row + 6}`);
                sheet.getRow(row + 6).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`E${row + 6}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${row + 6}`).border = { bottom: { style: 'thin' } };
                sheet.insertRow(row + 7, [
                    '', '', '', '', '', '(прописом)', '', '', '', '', '', '', ''
                ]);
                sheet.getRow(row + 7).font = { name: 'Times New Roman', size: 12, family: 4, italic: true };
                sheet.mergeCells(`E${row + 7}:M${row + 7}`);
                sheet.getCell(`E${row + 7}`).alignment = { vertical: 'top', horizontal: 'center' };

                //---------------------------------------//
                sheet.insertRow(row + 8, [
                    '',
                    '',
                    'загальна кількість одиниць за даними бухгалтерського обліку',
                    '', '', '', '',
                    `${Util.numberToDigit(value.amount)} ${Util.numberToString(value.amount, false)}`,
                    '', '', '', '', ''
                ]);
                sheet.mergeCells(`A${row + 8}:B${row + 8}`);
                sheet.mergeCells(`C${row + 8}:G${row + 8}`);
                sheet.mergeCells(`H${row + 8}:M${row + 8}`);
                sheet.getRow(row + 8).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`H${row + 8}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${row + 8}`).border = { bottom: { style: 'thin' } };
                sheet.insertRow(row + 9, [
                    '', '', '', '', '', '', '', '(прописом)', '', '', '', '', ''
                ]);
                sheet.getRow(row + 9).font = { name: 'Times New Roman', size: 12, family: 4, italic: true };
                sheet.mergeCells(`H${row + 9}:M${row + 9}`);
                sheet.getCell(`H${row + 9}`).alignment = { vertical: 'top', horizontal: 'center' };

                //---------------------------------------//
                sheet.insertRow(row + 10, [
                    '',
                    '',
                    'вартість за даними бухгалтерського обліку',
                    '', '',
                    `${Util.numberToDigit(value.startcost.toFixed(2))} ${Util.moneyToString(value.startcost)}`,
                    '', '', '', '', '', '', ''
                ]);
                sheet.mergeCells(`A${row + 10}:B${row + 10}`);
                sheet.mergeCells(`C${row + 10}:E${row + 10}`);
                sheet.mergeCells(`F${row + 10}:M${row + 10}`);
                sheet.getRow(row + 10).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`F${row + 10}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${row + 10}`).border = { bottom: { style: 'thin' } };
                sheet.insertRow(row + 11, [
                    '', '', '', '', '', '(прописом)', '', '', '', '', '', '', ''
                ]);
                sheet.getRow(row + 11).font = { name: 'Times New Roman', size: 12, family: 4, italic: true };
                sheet.mergeCells(`F${row + 11}:M${row + 11}`);
                sheet.getCell(`F${row + 11}`).alignment = { vertical: 'top', horizontal: 'center' };
                break;
            }
            case ActType.WITHOUTWEAR: {
                break;
            }
            case ActType.BABAK: {
                let row = sheet.lastRow != undefined ? sheet.lastRow.number + 1 : 1;
                sheet.insertRow(row,
                    ['Разом:',
                        '',
                        'x',
                        'x',
                        'x',
                        value.amount,
                        value.cost,
                        value.sum,
                        value.amount,
                        value.cost,
                        value.sum,
                        'x',
                    ]);
                sheet.mergeCells(`A${row}:B${row}`);
                sheet.getRow(row).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`C${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`D${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`E${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`F${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`G${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`H${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`I${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`J${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`K${row}`).alignment = { vertical: 'middle', horizontal: 'center' };
                sheet.getCell(`L${row}`).alignment = { vertical: 'middle', horizontal: 'center' };

                sheet.getCell(`A${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`C${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`D${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`E${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`F${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`G${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`H${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`I${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`J${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`K${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
                sheet.getCell(`L${row}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

                sheet.getCell(`G${row}`).numFmt = '#,##0.00';
                sheet.getCell(`H${row}`).numFmt = '#,##0.00';
                sheet.getCell(`J${row}`).numFmt = '#,##0.00';
                sheet.getCell(`K${row}`).numFmt = '#,##0.00';

                sheet.insertRow(row + 1, []);
                sheet.insertRow(row + 2, [
                    'Разом по сторінці:',
                    '',
                    'кількість порядкових номерів',
                    '',
                    `${Util.numberToDigit(value.count)} ${Util.numberToString(value.count, false)}`,
                    '', '', '', '', '', '', '', ''
                ]);
                sheet.mergeCells(`A${row + 2}:B${row + 2}`);
                sheet.mergeCells(`C${row + 2}:D${row + 2}`);
                sheet.mergeCells(`E${row + 2}:L${row + 2}`);
                sheet.getRow(row + 2).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`E${row + 2}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${row + 2}`).border = { bottom: { style: 'thin' } };
                sheet.insertRow(row + 3, [
                    '', '', '', '', '(прописом)', '', '', '', '', '', '', '', ''
                ]);
                sheet.getRow(row + 3).font = { name: 'Times New Roman', size: 12, family: 4, italic: true };
                sheet.mergeCells(`E${row + 3}:L${row + 3}`);
                sheet.getCell(`E${row + 3}`).alignment = { vertical: 'top', horizontal: 'center' };

                //---------------------------------------//
                sheet.insertRow(row + 4, [
                    '',
                    '',
                    'загальна кількість одиниць (фактично)',
                    '', '',
                    `${Util.numberToDigit(value.amount)} ${Util.numberToString(value.amount, false)}`,
                    '', '', '', '', '', '', ''
                ]);
                sheet.mergeCells(`A${row + 4}:B${row + 4}`);
                sheet.mergeCells(`C${row + 4}:E${row + 4}`);
                sheet.mergeCells(`F${row + 4}:L${row + 4}`);
                sheet.getRow(row + 4).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`F${row + 4}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${row + 4}`).border = { bottom: { style: 'thin' } };
                sheet.insertRow(row + 5, [
                    '', '', '', '', '', '(прописом)', '', '', '', '', '', '', ''
                ]);
                sheet.getRow(row + 5).font = { name: 'Times New Roman', size: 12, family: 4, italic: true };
                sheet.mergeCells(`F${row + 5}:L${row + 5}`);
                sheet.getCell(`F${row + 5}`).alignment = { vertical: 'top', horizontal: 'center' };

                //---------------------------------------//
                sheet.insertRow(row + 6, [
                    '',
                    '',
                    'сума фактична',
                    '',
                    `${Util.numberToDigit(value.sum.toFixed(2))} ${Util.moneyToString(value.sum)}`,
                    '', '', '', '', '', '', '', ''
                ]);
                sheet.mergeCells(`A${row + 6}:B${row + 6}`);
                sheet.mergeCells(`C${row + 6}:D${row + 6}`);
                sheet.mergeCells(`E${row + 6}:L${row + 6}`);
                sheet.getRow(row + 6).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`E${row + 6}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`E${row + 6}`).border = { bottom: { style: 'thin' } };
                sheet.insertRow(row + 7, [
                    '', '', '', '', '', '(прописом)', '', '', '', '', '', '', ''
                ]);
                sheet.getRow(row + 7).font = { name: 'Times New Roman', size: 12, family: 4, italic: true };
                sheet.mergeCells(`E${row + 7}:L${row + 7}`);
                sheet.getCell(`E${row + 7}`).alignment = { vertical: 'top', horizontal: 'center' };

                //---------------------------------------//
                sheet.insertRow(row + 8, [
                    '',
                    '',
                    'загальна кількість одиниць за даними бухгалтерського обліку',
                    '', '', '', '',
                    `${Util.numberToDigit(value.amount)} ${Util.numberToString(value.amount, false)}`,
                    '', '', '', '', ''
                ]);
                sheet.mergeCells(`A${row + 8}:B${row + 8}`);
                sheet.mergeCells(`C${row + 8}:G${row + 8}`);
                sheet.mergeCells(`H${row + 8}:L${row + 8}`);
                sheet.getRow(row + 8).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`H${row + 8}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`H${row + 8}`).border = { bottom: { style: 'thin' } };
                sheet.insertRow(row + 9, [
                    '', '', '', '', '', '', '', '(прописом)', '', '', '', '', ''
                ]);
                sheet.getRow(row + 9).font = { name: 'Times New Roman', size: 12, family: 4, italic: true };
                sheet.mergeCells(`H${row + 9}:L${row + 9}`);
                sheet.getCell(`H${row + 9}`).alignment = { vertical: 'top', horizontal: 'center' };

                //---------------------------------------//
                sheet.insertRow(row + 10, [
                    '',
                    '',
                    'сума за даними бухгалтерського обліку',
                    '', '',
                    `${Util.numberToDigit(value.sum.toFixed(2))} ${Util.moneyToString(value.sum)}`,
                    '', '', '', '', '', '', ''
                ]);
                sheet.mergeCells(`A${row + 10}:B${row + 10}`);
                sheet.mergeCells(`C${row + 10}:E${row + 10}`);
                sheet.mergeCells(`F${row + 10}:L${row + 10}`);
                sheet.getRow(row + 10).font = { name: 'Times New Roman', size: 12, family: 4, bold: true };
                sheet.getCell(`F${row + 10}`).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                sheet.getCell(`F${row + 10}`).border = { bottom: { style: 'thin' } };
                sheet.insertRow(row + 11, [
                    '', '', '', '', '', '(прописом)', '', '', '', '', '', '', ''
                ]);
                sheet.getRow(row + 11).font = { name: 'Times New Roman', size: 12, family: 4, italic: true };
                sheet.mergeCells(`F${row + 11}:L${row + 11}`);
                sheet.getCell(`F${row + 11}`).alignment = { vertical: 'top', horizontal: 'center' };
                break;
            }
            default: {
                console.log(`Unknown act type. From function 'insertFooter'`);
                break;
            }
        }
    }

    calculateRowHeight = (text, colWidth, defaultRowHeight) => {
        if (text.length <= colWidth) return defaultRowHeight;
        return Math.ceil(text.length / colWidth) * defaultRowHeight;
    }

    return {
        create: () => {
            let start = true;
            let temp = null;
            let isPredict = false;
            let tempIndex = null;

            total = null;

            assets.forEach((asset) => {
                // if(asset.name !='02') return;
                newWorkbook = new Excel.Workbook();
                sheet = newWorkbook.addWorksheet('Звіт', {
                    pageSetup: {
                        paperSize: 9,
                        orientation: 'landscape',
                        margins: {
                            left: 0.4,
                            top: 0.4,
                            right: 0.2,
                            bottom: 0.2,
                            footer: 0,
                            header: 0
                        },
                        scale: 60
                    }
                });

                height = 0;

                if (actType === ActType.WEAR) {
                    page = new Page(ActType.WEAR);
                } else if (actType === ActType.WITHOUTWEAR) {
                    page = new Page(ActType.WITHOUTWEAR);
                } else if (actType === ActType.BABAK) {
                    page = new Page(ActType.BABAK);
                }

                let ca = page.getArea();
    
                if (actType === ActType.WEAR) {
                    total = {
                        amount: 0,
                        startcost: 0,
                        endcost: 0,
                        wear: 0
                    }

                    let data = asset.getData();
                    data.forEach((value, index) => {
                        if (height >= ca.headerBegin && height <= ca.headerEnd) {
                            insertHeader(asset.type);
                            // if (start) {
                            //     sheet.getColumn(9).numFmt = '#,##0.00';
                            //     sheet.getColumn(12).numFmt = '#,##0.00';
                            //     sheet.getColumn(13).numFmt = '#,##0.00';
                            //     sheet.getColumn(14).numFmt = '#,##0.00';
                            //     start = false;
                            // }
                            height = ca.contentBegin;
                            console.log(`Page: ${page}`);
                        }
                        if (height >= ca.contentBegin && height <= ca.contentEnd) {
                            if (temp) {
                                // tempIndex=index;
                                insertRow([tempIndex, temp.name, temp.date, temp.code, '-', '-', 'шт.', 1, temp.startcost, '', 1, temp.startcost, temp.wear, temp.endcost, '', '']);
                                total.amount += 1;
                                total.startcost += temp.startcost;
                                total.wear += temp.wear;
                                total.endcost += temp.endcost;
                                height += calculateRowHeight(temp.name, sheet.getColumn(2).width - 7, sheet.properties.defaultRowHeight);
                                temp = null;
                                tempIndex = null;
                            }
                            isPredict = height + calculateRowHeight(value.name, sheet.getColumn(2).width - 7, sheet.properties.defaultRowHeight) <= ca.contentEnd;
                            if (isPredict) {
                                insertRow([index + 1, value.name, value.date, value.code, '-', '-', 'шт.', 1, value.startcost, '', 1, value.startcost, value.wear, value.endcost, '', '']);
                                total.amount += 1;
                                total.startcost += value.startcost;
                                total.wear += value.wear;
                                total.endcost += value.endcost;
                                height += calculateRowHeight(value.name, sheet.getColumn(2).width - 7, sheet.properties.defaultRowHeight);
                                isPredict = null;
                            } else {
                                temp = value;
                                tempIndex = index + 1;
                                height = ca.footerBegin;
                            }
                        }
                        if (height >= ca.footerBegin && height <= ca.footerEnd || data.length == index + 1) {
                            insertFooter(total);
                            height += page.getFooterHeight();
                            sheet.getRow(sheet.lastRow.number).addPageBreak();
                            page.next();
                            ca = page.getArea();
                            height = ca.headerBegin;
                            total.amount = 0;
                            total.startcost = 0;
                            total.wear = 0;
                            total.endcost = 0;
                            console.log(`footer height: ${height} `);
                        }
                    })
                    newWorkbook.xlsx.writeFile(`./reports/${asset.name}-ОЗ (знос).xlsx`);
                } else if (actType === ActType.WITHOUTWEAR) {

                } else if (actType === ActType.BABAK) {
                    total = {
                        count: 0,
                        amount: 0,
                        cost: 0,
                        sum: 0
                    }

                    let data = asset.getData();
                    data.forEach((value, index) => {
                        if (height >= ca.headerBegin && height <= ca.headerEnd) {
                            insertHeader();
                            height = ca.contentBegin;
                            console.log(`Page: ${ca.page}, Header height:${height}`);
                        }
                        if (height >= ca.contentBegin && height <= ca.contentEnd) {
                            if (temp) {
                                // tempIndex=index;
                                insertRow([tempIndex, temp.bill, temp.name, '-', temp.unit, temp.amount, temp.cost, temp.sum, temp.amount, temp.cost, temp.sum, '']);
                                total.count += 1;
                                total.amount += temp.amount;
                                total.cost += temp.cost;
                                total.sum += temp.sum;
                                height += calculateRowHeight(temp.name, sheet.getColumn(2).width - 7, sheet.properties.defaultRowHeight);
                                temp = null;
                                tempIndex = null;
                            }
                            isPredict = height + calculateRowHeight(value.name, sheet.getColumn(2).width - 7, sheet.properties.defaultRowHeight) <= ca.contentEnd;
                            if (isPredict) {
                                insertRow([index + 1, value.bill, value.name, '-', value.unit, value.amount, value.cost, value.sum, value.amount, value.cost, value.sum, '']);
                                total.count += 1;
                                total.amount += value.amount;
                                total.cost += value.cost;
                                total.sum += value.sum;
                                height += calculateRowHeight(value.name, sheet.getColumn(2).width - 7, sheet.properties.defaultRowHeight);
                                isPredict = null;
                            } else {
                                temp = value;
                                tempIndex = index + 1;
                                height = ca.footerBegin;
                            }
                        }
                        if (height >= ca.footerBegin && height <= ca.footerEnd || data.length == index + 1) {
                            insertFooter(total);
                            height += page.getFooterHeight();
                            sheet.getRow(sheet.lastRow.number).addPageBreak();
                            page.next();
                            ca = page.getArea();
                            height = ca.headerBegin;
                            total.count = 0;
                            total.amount = 0;
                            total.cost = 0;
                            total.sum = 0;
                            temp=null;
                            tempIndex = null;
                            console.log(`footer height: ${height} `);
                        }
                    })
                    newWorkbook.xlsx.writeFile(`./reports/${asset.name}-ОЗ (Бабак).xlsx`);
                }
            })
            return assets;
        },
        createTitle: async (actType, asset, subjects) => {
            const getSubject = (subjects, asset) => {
                let title = subjects.find(item => item.code == asset.name);
                return title === undefined ? 'Рахунок у довіднику відсутній' : title.name;
            }

            switch (actType) {
                case ActType.WITHOUTWEAR: {
                    let workbook = new Excel.Workbook();
                    workbook.xlsx.readFile('./template/TEMPLATE_FIXED.xlsx')
                        .then(() => {
                            //---------------titel sheet---------------------------//
                            let titleSheet = workbook.getWorksheet('Титулка');
                            let cell = titleSheet.getCell('C18:H18');
                            cell.value = getSubject(subjects, asset);

                            //---------------total sheet---------------------------//
                            let totalSheet = workbook.getWorksheet('Підсумок');
                            let cellCount = totalSheet.getCell('I1:T1');
                            let cellStartAmount = totalSheet.getCell('J3:T3');
                            let cellStartCost = totalSheet.getCell('M5:T5');
                            let cellEndAmount = totalSheet.getCell('M7:T7');
                            let cellEndCost = totalSheet.getCell('K9:T9');

                            let total = asset.getTotal();
                            cellCount.value = `${Util.numberToDigit(total.count)} ${Util.numberToString(total.count)}`;
                            cellStartAmount.value = `${Util.numberToDigit(total.amount)} ${Util.numberToString(total.amount)}`;
                            cellStartCost.value = `${Util.numberToDigit(total.startcost.toFixed(2))} ${Util.moneyToString(total.startcost)}`;
                            cellEndAmount.value = `${Util.numberToDigit(total.amount)} ${Util.numberToString(total.amount)}`;
                            cellEndCost.value = `${Util.numberToDigit(total.endcost.toFixed(2))} ${Util.moneyToString(total.endcost)}`;

                            return workbook.xlsx.writeFile(`./reports/${asset.name}-ОЗ Титулка.xlsx`);
                        }).then(() => {
                            console.log('write file ' + asset.name);
                        }).catch(error => console.log(error));
                    break;
                }
                case ActType.WEAR: {
                    let workbook = new Excel.Workbook();
                    workbook.xlsx.readFile('./template/TEMPLATE_FIXED.xlsx')
                        .then(() => {
                            //---------------titel sheet---------------------------//
                            let titleSheet = workbook.getWorksheet('Титулка');
                            let cell = titleSheet.getCell('C18:H18');
                            cell.value = getSubject(subjects, asset);

                            //---------------total sheet---------------------------//
                            let totalSheet = workbook.getWorksheet('Підсумок');
                            let cellCount = totalSheet.getCell('I1:T1');
                            let cellStartAmount = totalSheet.getCell('J3:T3');
                            let cellStartCost = totalSheet.getCell('M5:T5');
                            let cellEndAmount = totalSheet.getCell('M7:T7');
                            let cellEndCost = totalSheet.getCell('K9:T9');

                            let total = asset.getTotal();
                            cellCount.value = `${Util.numberToDigit(total.count)} ${Util.numberToString(total.count)}`;
                            cellStartAmount.value = `${Util.numberToDigit(total.amount)} ${Util.numberToString(total.amount)}`;
                            cellStartCost.value = `${Util.numberToDigit(total.startcost.toFixed(2))} ${Util.moneyToString(total.startcost)}`;
                            cellEndAmount.value = `${Util.numberToDigit(total.amount)} ${Util.numberToString(total.amount)}`;
                            cellEndCost.value = `${Util.numberToDigit(total.endcost.toFixed(2))} ${Util.moneyToString(total.endcost)}`;

                            return workbook.xlsx.writeFile(`./reports/${asset.name} - ОЗ Титулка (знос).xlsx`);
                        }).then(() => {
                            console.log('write file ' + asset.name);
                        }).catch(error => console.log(error));
                    break;
                }
                case ActType.BABAK: {
                    let workbook = new Excel.Workbook();
                    workbook.xlsx.readFile('./template/TEMPLATE_INTANGIBLE.xlsx')
                        .then(() => {
                            //---------------titel sheet---------------------------//
                            let titleSheet = workbook.getWorksheet('Титулка');
                            let cell = titleSheet.getCell('C18:H18');
                            cell.value = getSubject(subjects, asset);

                            //---------------total sheet---------------------------//
                            let totalSheet = workbook.getWorksheet('Підсумок');
                            let cellCount = totalSheet.getCell('I1:T1');
                            let cellStartAmount = totalSheet.getCell('J3:T3');
                            let cellStartSum = totalSheet.getCell('M5:T5');
                            let cellEndAmount = totalSheet.getCell('M7:T7');
                            let cellEndSum = totalSheet.getCell('K9:T9');

                            let total = asset.getTotal();
                            cellCount.value = `${Util.numberToDigit(total.count)} ${Util.numberToString(total.count)}`;
                            cellStartAmount.value = `${Util.numberToDigit(total.amount)} ${Util.numberToString(total.amount)}`;
                            cellStartSum.value = `${Util.numberToDigit(total.sum.toFixed(2))} ${Util.moneyToString(total.sum)}`;
                            cellEndAmount.value = `${Util.numberToDigit(total.amount)} ${Util.numberToString(total.amount)}`;
                            cellEndSum.value = `${Util.numberToDigit(total.sum.toFixed(2))} ${Util.moneyToString(total.sum)}`;

                            return workbook.xlsx.writeFile(`./reports/${asset.name} - ОЗ Титулка (БАБАК).xlsx`);
                        }).then(() => {
                            console.log('write file ' + asset.name);
                        }).catch(error => console.log(error));
                    break;
                }
                default:
                    console.log(`Unknown act type. From function 'create Title'`);
                    break;
            }
        },
    }
}

module.exports = {
    init
}