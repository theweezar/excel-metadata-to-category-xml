'use strict';

const path = require('path');
const fs = require('fs');
const excelToJson = require('convert-excel-to-json');
const xml2js = require('xml2js');

// Set attribute key name
const configParse = {
    attrkey: 'attribute'
};

// Set attribute key name
const configBuild = {
    attrkey: 'attribute'
};

const jsToXmlFile = function (fileName, obj, callBack) {
    var filepath = path.normalize(path.join(__dirname, fileName));
    var builder = new xml2js.Builder(configBuild);
    var xml = builder.buildObject(obj);
    fs.writeFile(filepath, xml, callBack);
}

const processCatalog = function (xmlObj) {
    var catalog = xmlObj.catalog;
    var category = catalog.category;

    if (catalog['category-assignment']) {
        delete catalog['category-assignment'];
    }

    console.log(category);

    jsToXmlFile(
        path.join('result-xml', 'result.xml'),
        xmlObj,
        () => {
            console.log('Done');
        }
    );
}

const afterParsingXML = function (error, result) {
    if (error) {
        console.log(error);
    } else if (result) {
        processCatalog(result);
    }
}

const buildMetaDataObject = function (row) {
    var title = row.C;
    var desc = row.D;
    var keyword = row.E;
    return {
        categoryId: row.B,
        xmlString: `
        <page-attributes>
            <page-title xml:lang="x-default">${title}</page-title>
            <page-description xml:lang="x-default">${desc}</page-description>
            <page-keywords xml:lang="x-default">${keyword}</page-keywords>
            <page-url xml:lang="x-default">luggage</page-url>
        </page-attributes>
        `
    };
}

const processXLSX = function (xlsxObject) {
    var metaTags = xlsxObject['Meta Tags'];
    var metaDataObjectList = [];
    // console.log(metaTags);

    metaTags.forEach(row => {
        if (typeof row === 'object' && Object.keys(row).length === 5) {
            var metaDataObject = buildMetaDataObject(row);
            metaDataObjectList.push(metaDataObject);
        }
    });

    console.log(metaDataObjectList);
}

const main = function () {
    const xlsxObject = excelToJson({
        sourceFile: path.join(__dirname, 'excel', 'AT_KR_SEO_Redirections_v1.0.xlsx')
    });

    processXLSX(xlsxObject)

    // const parser = new xml2js.Parser(configParse);

    // const xmlString = fs.readFileSync(path.join(__dirname, 'raw-xml', '20220429_catalog_kr-americantourister.xml'), "utf8");

    // parser.parseString(xmlString, afterParsingXML);
}

main();