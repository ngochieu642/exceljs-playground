const am = require('am');
const ExcelJS = require('exceljs');
const _ = require('lodash');
const fs = require('fs');

const sheetOptions = {
  mainSheet: {
    name: 'Product Plan',
    options: {
      header: [
        'machine',
        'engine',
        'dev',
        'product_name',
        'process',
        'standard_motor_working_time',
        'standard_standby_time',
        'quantity_produced',
        'start_time',
        'start_day',
        'sip_device_id',
      ],
    },
  },
  mappingSheet: {
    name: 'Mapping',
    options: {
      header: [
        'client_name',
        'engine',
        'dev',
        'sip_device_id',
        'mac_address',
        'device_local_id',
      ],
    },
  },
};

const writeWorkbook = async (filename, toWriteWorkbook) => {
  await toWriteWorkbook.xlsx.writeFile(filename)
}

class ExcelJSHeader {
  constructor(header, key, width=null) {
    this.header = header;
    this.key = key;

    if (!width) {
      this.width = this.header.length * 2;
    } else {
      this.width = width
    }
  }

  get excelJSFormat() {
    return {
      header: this.header,
      key: this.key,
      width: this.width,
    }
  }
}

class MappingRowObject {
  constructor({client_name, engine, sip_device_id, mac_address, device_local_id}) {
    this.customDeviceName = client_name;
    this.deviceName = engine;
    this.sipDeviceId = sip_device_id;
    this.macAddress = mac_address;
    this.deviceLocalId = device_local_id;
  }

  get propertiesList() {
    return Object.keys(this);
  }

  get baseProps() {
    return {
      customDeviceName: this.customDeviceName || null,
      deviceName: this.deviceName || null,
      sipDeviceId: this.sipDeviceId || null,
      macAddress: this.macAddress || null,
      deviceLocalId: this.deviceLocalId || null,
    }
  }
}

const alphabetArray = [...Array(26)].map((_, i) =>
  String.fromCharCode('A'.charCodeAt(0) + i)
);

const alphabetNumberMapping = alphabetArray.reduce((accumulator, nextLetter, index) => {
  return {...accumulator, [nextLetter]: index + 1};
}, {})

const numberAlphabetMapping = alphabetArray.reduce((accumulator, nextLetter, index) => {
  return {...accumulator, [index + 1]: nextLetter};
}, {})

const alphabetToNumber = (alphabet) => {
  return alphabetNumberMapping[alphabet];
}

const numberToAlphabet = (number) => {
  return numberAlphabetMapping[number];
}

const addMappingSheet = (workbook) => {
  // Data to write
  let {data} = JSON.parse(fs.readFileSync('./Mapping.json', {encoding: 'utf-8'}))

  let rowObjects = data.map(x => {
    return new MappingRowObject(x);
  })

  // Force workbook calculation onload
  workbook.calcProperties.fullCalcOnLoad = true;

  const mappingSheet = workbook.addWorksheet('Mapping', {properties: {tabColor: {argb: '00ad9d9d'}}});

  // Columns
  mappingSheet.columns = [
    {header: 'machine', key: 'customDeviceName', width: 30},
    {header: 'engine', key: 'deviceName', width: 30},
    {header: 'dev', key: 'devKey', width: 20},
    {header: 'sip_device_id', key: 'sipDeviceId', width: 46},
    {header: 'mac_address', key: 'macAddress', width: 20},
    {header: 'device_local_id', key: 'deviceLocalId', width: 20},
  ]

  // Hide the dev col
  let devCol = mappingSheet.getColumn('devKey');
  devCol.hidden = false;

  // Rows

  // Add objects from data to sheet
  rowObjects.forEach(object => {
    mappingSheet.addRow(object);
  })

  // Add CONCATENATE(A1, A2)
  let customDeviceCol = mappingSheet.getColumn('customDeviceName');
  let customDeviceColAlphabet = numberToAlphabet(customDeviceCol._number);

  let deviceCol = mappingSheet.getColumn('deviceName');
  let deviceColAlphabet = numberToAlphabet(deviceCol._number);

  mappingSheet.eachRow({includeEmpty: false}, function (row, rowNumber) {
    // Header row
    if (rowNumber === 1) {
      return;
    }

    // Else
    let devCell = row.getCell('devKey');
    devCell.value = {
      formula: `CONCATENATE(${customDeviceColAlphabet}${rowNumber},${deviceColAlphabet}${rowNumber})`
    }
  });

  return workbook;
}

const addProductPlanSheet = (workbook, allowedRows) => {
  const productPlanSheet = workbook.addWorksheet('Product Plan', {properties: {tabColor: {argb: '00d9adad'}}})

  productPlanSheet.columns = [
    {header: 'machine', key: 'customDeviceName', width: 30},
    {header: 'engine', key: 'deviceName', width: 30},
    {header: 'dev', key: 'devKey', width: 20},
    {header: 'product_name', key: 'productName', width: 20},
    {header: 'process', key: 'processName', width: 20},
    {header: 'standard_motor_working_time', key: 'motorWorkingTime', width: 30},
    {header: 'standard_standby_time', key: 'standbyTime', width: 30},
    {header: 'quantity_produced', key: 'quantityProduced', width: 20},
    {header: 'start_time', key: 'startTime', width: 20},
    {header: 'start_date', key: 'startDate', width: 20},
    {header: 'sip_device_id', key: 'sipDeviceId', width: 46},
  ];

  // Last row to extend sheet range
  const lastRow = productPlanSheet.getRow(allowedRows + 1); // Since we have had header row

  // Add CONCATENATE & VLOOKUP
  let devCol = productPlanSheet.getColumn('devKey');
  let devColAlphabet = numberToAlphabet(devCol._number);

  let customDeviceCol = productPlanSheet.getColumn('customDeviceName');
  let customDeviceColAlphabet = numberToAlphabet(customDeviceCol._number);

  let deviceCol = productPlanSheet.getColumn('deviceName');
  let deviceColAlphabet = numberToAlphabet(deviceCol._number);

  productPlanSheet.eachRow({includeEmpty: true}, function (row, rowNumber) {
    // Header row
    if (rowNumber === 1) {
      return;
    }

    // Else
    let devCell = row.getCell('devKey');
    devCell.value = {
      formula: `CONCATENATE(${customDeviceColAlphabet}${rowNumber},${deviceColAlphabet}${rowNumber})`
    }

    let sipDeviceIdCell = row.getCell('sipDeviceId');
    sipDeviceIdCell.value = {
      formula: `VLOOKUP(${devColAlphabet}${rowNumber}, ${sheetOptions.mappingSheet.name}!$C:$F, 2, 0)`,
    }
  });

  return workbook;
}

const validName = (inputName) => {
  return inputName.toLowerCase().trim().replace(/\s\s+/, '');
};

const addCustomDeviceSheet = (workbook) => {
  const customDeviceSheet = workbook.addWorksheet('Custom Device', {properties: {tabColor: {argb: '00f5a25d'}}})

  // Achieve custom device listr
  let {data} = JSON.parse(fs.readFileSync('./Mapping.json', {encoding: 'utf-8'}))
  let customDeviceNameUniqArr = _.uniq(data.map(x => {
    try {
      return x['client_name'];
    } catch (e) {
      return null;
    }
  }))
  customDeviceNameUniqArr = customDeviceNameUniqArr.filter(x => !!x);

  // Add 1 CustomDevice Column & many specific-column
  let columnArr = [
    {header: 'Custom Device', key: 'customDevice', width: 30}
  ]

  let generatedHeaders = customDeviceNameUniqArr.map(x => {
    let headerObj = new ExcelJSHeader(x, x);
    return headerObj.excelJSFormat;
  })

  columnArr = [...columnArr, ...generatedHeaders];

  customDeviceSheet.columns = columnArr;

  // Custom Device Column -> Fill in all Custom Device Name, Unique
  let customDeviceCol = customDeviceSheet.getColumn('customDevice');
  customDeviceCol.values = [customDeviceCol.values[1], ...customDeviceNameUniqArr]; // customDeviceCol.values = [undefined, 'header name']

  // Define Cell Range from A2 -> A(end) to become
  customDeviceCol.eachCell({includeEmpty: false}, function (cell, rowNumber) {
    console.log(rowNumber);
    cell.name = 'customdevice';
  })

  // For each Custom Device
    // Add a column name = Custom Device Name
    // Define cell Range from colAlphabet (2) -> colAlphabet(end) to become validName(customDeviceName)
  return workbook;
}

const addListLimitToProductPlanSheet = (workbook) => {
  // For customDevice column, limit to 'customdevice'
  const productPlanSheet = workbook.getWorksheet('Product Plan');
  const lastRow = productPlanSheet.lastRow;

  let customDeviceCol = productPlanSheet.getColumn('customDeviceName');
  let customDeviceColAlphabet = numberToAlphabet(customDeviceCol._number);

  productPlanSheet.eachRow({includeEmpty: true}, function(row, rowNumber) {
    // Heade row
    if (rowNumber === 1) {
      return;
    }

    // Else
    let customDeviceCell = row.getCell('customDeviceName');
    customDeviceCell.dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: '=customdevice'
    }
  })

  // For device column, limit = INDIRECT(SUBSTITUTE(A$, " ", ""))
  return workbook;
}

async function main() {
  let workbook = new ExcelJS.Workbook();

  // Add MAPPING SHEET
  workbook = addMappingSheet(workbook);

  // Add PRODUCT PLAN SHEET
  workbook = addProductPlanSheet(workbook, 100);

  // Add Custom Device & Device Sheet (for limit range)
  workbook = addCustomDeviceSheet(workbook);

  workbook = addListLimitToProductPlanSheet(workbook);

  // Write file
  let filename = './hieu.xlsx';
  await writeWorkbook(filename, workbook);
  console.log('Done')
}

am(main);