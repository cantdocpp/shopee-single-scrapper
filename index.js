'use strict'

require('events').EventEmitter.defaultMaxListeners = 1000;
require('dotenv').config()
const fs = require('fs')
const cheerio = require('cheerio');
const puppeteer = require('puppeteer');
const cloudinary = require('cloudinary').v2
const { ws, wb } = require('./excel');

// Get this at your cloudinary dashboard
cloudinary.config({
    cloud_name: process.env.CLOUD_NAME,
    api_key: process.env.API_KEY,
    api_secret: process.env.API_SECRET
})

const getHtml = async (url) => {
    console.log('start scrape from the url page...');
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(url, { waitUntil: 'networkidle2' });
    const html = await page.evaluate(() => document.querySelector('*').outerHTML);

    return html;
}

const loadHtml = (html) => {
    return cheerio.load(html);
}

const getCleanImageLink = (link) => {
    const splitStyle = link.split(';');
    const splitUrl = splitStyle[0].split(':');
    const cleanUrl = splitUrl[2].slice(0, -5);
    const fullUrl = 'https:' + cleanUrl;

    return fullUrl;
}

const generateData = async ($) => {
    const sizes = [];
    const colors = [];
    const imgs = [];
    
    const category = 18200;
    const name = $('.qaNIZv span').text();
    const price = $('._3n5NQx').text();
    const description = $('._2u0jt9 span').text();
    const stock = 1000;
    const weight = 100;

    await $('.product-variation').each((index, element) => {
        const elementItem = $(element).text();
        if (index <= 3) {
            sizes.push(elementItem);
        } else if (index > 3) {
            colors.push(elementItem);
        }
    })

    await $('._2Fw7Qu').each(async (index, element) => {
        const link = $(element).attr('style');
        const imageLink = await getCleanImageLink(link);
        imgs.push(imageLink);
    });

    console.log(price, '........ price')

    const excelDataObject = {};
    excelDataObject.name = name;
    excelDataObject.category = category;
    excelDataObject.priceMin = price.split(' - ')[0];
    excelDataObject.priceMiddle = 50000;
    excelDataObject.priceMax = 55000;
    excelDataObject.description = description;
    excelDataObject.stock = stock;
    excelDataObject.weight = weight
    excelDataObject.imgs = imgs;
    excelDataObject.sizes = sizes;
    excelDataObject.colors = colors;

    return excelDataObject;
}

const readImage = () => {
    let listImageName = fs.readdirSync('./images/');
    
    return listImageName;
}

const getImageLink = async (image) => {
    console.log(image, 'image')
    try {
        return new Promise((resolve, reject) => {
            cloudinary.uploader.upload(`./images/${image + 1}.jpg`, {
                use_filename: true,
                unique_filename: false
            } ,(err, url) => {
              if (err) return reject(err);
              return resolve(url);
            })
        });
    } catch(error) {
        console.log(error)
    }
}

const addEmptyColumn = (column, ws) => {
    ws.cell(column, 1)
        .string('')

    ws.cell(column, 2)
        .string('')

    ws.cell(column, 3)
        .string('')

    ws.cell(column, 4)
        .string('')

    ws.cell(column, 5)
        .string('')

    ws.cell(column, 6)
        .string('')

    ws.cell(column, 7)
        .string('')

    ws.cell(column, 8)
        .string('')

    ws.cell(column, 9)
        .string('')

    ws.cell(column, 10)
        .string('')

    ws.cell(column, 11)
        .string('')

    ws.cell(column, 12)
        .string('')

    ws.cell(column, 13)
        .string('')

    ws.cell(column, 14)
        .string('')

    ws.cell(column, 15)
        .string('')

    ws.cell(column, 16)
        .string('')

    ws.cell(column, 17)
        .string('')

    ws.cell(column, 18)
        .string('')

    ws.cell(column, 19)
        .string('')

    ws.cell(column, 20)
        .string('')

    ws.cell(column, 21)
        .string('')

    ws.cell(column, 22)
        .string('')

    ws.cell(column, 23)
        .string('')

    ws.cell(column, 24)
        .string('')

    ws.cell(column, 25)
        .string('')

    ws.cell(column, 26)
        .string('')

    ws.cell(column, 27)
        .string('')

    ws.cell(column, 28)
        .string('')

    ws.cell(column, 29)
        .string('')

    ws.cell(column, 30)
        .string('')

    ws.cell(column, 31)
        .string('')

    ws.cell(column, 32)
        .string('')

    ws.cell(column, 33)
        .string('')

    ws.cell(column, 34)
        .string('')

    ws.cell(column, 35)
        .string('')
}

const writeVarianceToExcel = (ws, excelDataObject, column, varianceIntegradeCode, mainImage) => {
    return new Promise((resolve, reject) => {
        for (let k = 0; k < excelDataObject.colors.length; k++) {
            for (let l = 0; l < excelDataObject.sizes.length; l++) {
                ws.cell(column, 1)
                    .number(excelDataObject.category)
                ws.cell(column, 2)
                    .string(excelDataObject.name)
                ws.cell(column, 3)
                    .string(excelDataObject.description)
                ws.cell(column, 5)
                    .number(varianceIntegradeCode)
                ws.cell(column, 6)
                    .string('ukuran')
                ws.cell(column, 7)
                    .string(excelDataObject.sizes[l])

                ws.cell(column, 9)
                    .string('warna')
                ws.cell(column, 10)
                    .string(excelDataObject.colors[k])
                
                if (l === 0 || l === 1) {
                    ws.cell(column, 11)
                        .string(excelDataObject.priceMin)
                } else if (l === 2) {
                    ws.cell(column, 11)
                        .number(excelDataObject.priceMiddle)
                } else if (l === 3) {
                    ws.cell(column, 11)
                        .number(excelDataObject.priceMax)
                }

                ws.cell(column, 12)
                    .number(excelDataObject.stock)
                
                console.log(l, 'sizes index')
                ws.cell(column, 14)
                    .string(mainImage)
                
                ws.cell(column, 15)
                    .string(excelDataObject.imgs[1])
                ws.cell(column, 16)
                    .string(excelDataObject.imgs[2])

                ws.cell(column, 23)
                    .number(excelDataObject.weight)
                
                column++;
            }
        }
        resolve();
    })
}


const writeExcel = async (excelDataObject) => {
    const listImageName = readImage();
    const varianceQuantity = excelDataObject.sizes.length * excelDataObject.colors.length;
    let varianceCodeBase = 100;
    let column = 2;
    
    for (let i = 0; i <= 45; i++) {
        const varianceIntegradeCode = varianceCodeBase + i;
        await getImageLink(i)
            .then( (response) => {
                for (let j = 0; j < varianceQuantity; j++) {
                    const mainImage = response.url;
    
                    // ws.cell(column, 1)
                    //     .number(excelDataObject.category)
                    // ws.cell(i + 1, 2)
                    //     .string(excelDataObject.name)
                    // ws.cell(i + 1, 3)
                    //     .string(excelDataObject.description)
                    // ws.cell(i + 1, 5)
                    //     .number(varianceIntegradeCode)

                    writeVarianceToExcel(ws, excelDataObject, column, varianceIntegradeCode, mainImage)
                        .then(async (res) => {
                            console.log('write')
                            await wb.write('upload.xlsx');
                            column++;
                            console.log(column)
                        })
                        .catch(error => {
                            console.log(error)
                        })
    
                    // ws.cell(i + 1, 11)
                    //     .string(excelDataObject.price)
                    // ws.cell(i + 1, 12)
                    //     .number(excelDataObject.stock)
    
                    // ws.cell(i + 1, 14)
                    //     .string(mainImage)
                    
                    // ws.cell(i + 1, 15)
                    //     .string(excelDataObject.imgs[1])
                    // ws.cell(i + 1, 16)
                    //     .string(excelDataObject.imgs[2])
    
                    // ws.cell(i + 1, 23)
                    //     .number(excelDataObject.weight)
    
                    // await wb.write('upload.xlsx');
                    // column++;
                    // console.log(column)
                }
            })
    }
}

const app = async () => {
    const url = 'https://shopee.co.id/Kaos-Champion-Pria-Distro-PREMIUM-QUALITY-i.226091.4433868876';
    const html = await getHtml(url);
    const $ = await loadHtml(html);
    const excelDataObject = await generateData($);
    writeExcel(excelDataObject);
}

// Run the app
app();
