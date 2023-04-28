require('dotenv').config();
const { Opengram, session } = require('opengram');
const axios = require('axios');
const fs = require('fs');
const readXlsx = require('read-excel-file/node');
const mysql = require('mysql2/promise');



const bot = new Opengram(process.env.BOT_TOKEN);

bot.use(Opengram.log())
bot.use(session())

bot.start(async (ctx) => {
    await ctx.replyWithHTML(`\n\nДля беспроблемной работы используйте готовые шаблонные файлы \n\n* <code>product.xlsx</code> \n\n* <code>manufacturer.xlsx</code> \n\n* <code>category.xlsx</code> \n\nВызвать шаблоны можно командами \n\n/tProd \n\n/tMan \n\n/tCat\n\n/manufacturers - для просмотра производителей\n\n/categories - для просмотра категорий\n\n/prodImport - для добавления товаров после загрузки <code>product.xlsx</code>\n\n/manImport - для добавления производителей после загрузки <code>manufacturer.xlsx</code>\n\n/catImport - для добавления категорий после загрузки <code>category.xlsx</code>`)
})

bot.command('tProd', async ctx => {
    const xlsxPath = './static/product.xlsx'
    await ctx.replyWithDocument({source: fs.createReadStream(xlsxPath), filename: 'product.xlsx'})
})

bot.command('tCat', async ctx => {
    const xlsxPath = './static/category.xlsx'
    await ctx.replyWithDocument({source: fs.createReadStream(xlsxPath), filename: 'category.xlsx'})
})

bot.command('tMan', async ctx => {
    const xlsxPath = './static/manufacturer.xlsx'
    await ctx.replyWithDocument({source: fs.createReadStream(xlsxPath), filename: 'manufacturer.xlsx'})
})

bot.command('manufacturers', async ctx => {

    const mSql = `SELECT name, manufacturer_id FROM oc_manufacturer`

    await mysql.createConnection({host: '127.0.0.1', user: 'brut', password: 'brut', database: 'ocdb'})
        .then(con => con.query(mSql))
        .then((rows, fields) => {
            const fe = rows[0]
            fe.forEach(async element => await ctx.reply(element))
        })
        .catch(error => ctx.reply(error));
})

bot.command('categories', async ctx => {

    const cSql = `SELECT name, category_id FROM oc_category_description`

    await mysql.createConnection({host: '127.0.0.1', user: 'brut', password: 'brut', database: 'ocdb'})
        .then(con => con.query(cSql))
        .then((rows, fields) => {
            const fe = rows[0]
            fe.forEach(async element => await ctx.reply(element))
        })
        .catch(error => ctx.reply(error));
})

bot.command('manImport', async ctx => {
    const title = ctx.session.number
    if (title) {
        await fs.stat(`C:/Users/Adam/Taxi_bot/static/${title}`, async function(err, stats) {
            if (err) {
                throw err;
            } else {
               await readXlsx(`C:/Users/Adam/Taxi_bot/static/${title}`).then(async(rows) => {
    
                    // ok it works
                    const datas = rows.slice(1)
                    const header = rows.slice(0, 1).flat()
    
                    const result = {};
    
                    for (let i = 0; i < datas.length; i++) {
                        const obj = {};
    
                        for (let j = 0; j < header.length; j++) {
                            obj[header[j]] = datas[i][j];
                        }
                        
                        result[i + 1] = obj;
                    }
                    
                    for (let key in result) {
                        const element = result[key];
                        const catSql =  `BEGIN; INSERT INTO oc_manufacturer(name, image, sort_order) VALUES(?, ?, ?); SET @manufacturer_id=LAST_INSERT_ID(); INSERT INTO oc_manufacturer_to_layout(manufacturer_id, store_id, layout_id) VALUES(@manufacturer_id, ?, ?); INSERT INTO oc_manufacturer_to_store(manufacturer_id, store_id) VALUES(@manufacturer_id, ?); COMMIT;`
                        
                        await mysql.createConnection({host: '127.0.0.1', user: 'brut', password: 'brut', database: 'ocdb', multipleStatements: true})
                            .then(con => con.query(catSql, [ element.name, element.image, element.sort_order, element.store_id, element.layout_id, element.store_id ]))
                            .then(ctx.reply('Производитель добавлен'))
                            .catch(error => ctx.reply(error));                    
                    };

                })  
            }
        })
    } else {
        await ctx.replyWithHTML('Вы видите эту ошибку, потому что не отправили .xlsx файл или время сессии истекло\n\nПожалуйста, снова отправьте боту .xlsx файл и сразу после загрузки отправьте команду /manImport')
    }
})

bot.command('catImport', async ctx => {
    const title = ctx.session.number
    if (title) {
        await fs.stat(`C:/Users/Adam/Taxi_bot/static/${title}`, async function(err, stats) {
            if (err) {
                throw err;
            } else {
               await readXlsx(`C:/Users/Adam/Taxi_bot/static/${title}`).then(async(rows) => {
    
                    // ok it works
                    const datas = rows.slice(1)
                    const header = rows.slice(0, 1).flat()
    
                    const result = {};
    
                    for (let i = 0; i < datas.length; i++) {
                        const obj = {};
    
                        for (let j = 0; j < header.length; j++) {
                            obj[header[j]] = datas[i][j];
                        }
                        
                        result[i + 1] = obj;
                    }
                    
                    for (let key in result) {
                        const element = result[key];
                        const catSql =  `BEGIN; SET @date=NOW(); INSERT INTO oc_category(image, parent_id, top, \`column\`, sort_order, status, date_added, date_modified) VALUES(?, ?, ?, ?, ?, ?, @date, @date); SET @category_id=LAST_INSERT_ID(); INSERT INTO oc_category_description(category_id, language_id, name, description, meta_title, meta_description, meta_keyword) VALUES(@category_id, ?, ?, ?, ?, ?, ?); INSERT INTO oc_category_path(category_id, path_id, level) VALUES(@category_id, @category_id, ?); INSERT INTO oc_category_to_layout(category_id, store_id, layout_id) VALUES(@category_id, ?, ?); INSERT INTO oc_category_to_store(category_id, store_id) VALUES(@category_id, ?); COMMIT;`
                        
                        await mysql.createConnection({host: '127.0.0.1', user: 'brut', password: 'brut', database: 'ocdb', multipleStatements: true})
                            .then(con => con.query(catSql, [element.image, element.parent_id, element.top, element.column, element.sort_order, element.status, element.language_id, element.name, element.description, element.meta_title, element.meta_description, element.meta_keyword, element.level, element.store_id, element.layout_id, element.store_id ]))
                            .then(ctx.reply('Категория добавлена'))
                            .catch(error => ctx.reply(error));                    
                      };
                })
            }
        })
    } else {
        await ctx.replyWithHTML('Вы видите эту ошибку, потому что не отправили .xlsx файл или время сессии истекло\n\nПожалуйста, снова отправьте боту .xlsx файл и сразу после загрузки отправьте команду /catImport')
    }
})

bot.command('prodImport', async ctx => {
    const title = ctx.session.number
    if (title) {
        await fs.stat(`C:/Users/Adam/Taxi_bot/static/${title}`, async function(err, stats) {
            if (err) {
                throw err;
            } else {
               await readXlsx(`C:/Users/Adam/Taxi_bot/static/${title}`).then(async(rows) => {
    
                    // ok it works
                    const datas = rows.slice(1)
                    const header = rows.slice(0, 1).flat()
    
                    const result = {};
    
                    for (let i = 0; i < datas.length; i++) {
                        const obj = {};
    
                        for (let j = 0; j < header.length; j++) {
                            obj[header[j]] = datas[i][j];
                        }
                        
                        result[i + 1] = obj;
                    }
                    
                    for (let key in result) {
                        const element = result[key];
                        const iSql =  `BEGIN; SET @date=NOW(); INSERT INTO oc_product(master_id, model, sku, upc, ean, jan, isbn, mpn, location, variant, override, quantity, minimum, subtract, stock_status_id, date_available, manufacturer_id, shipping, price, points, weight, weight_class_id, length, width, height, length_class_id, status, tax_class_id, sort_order, date_added, date_modified, image) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, @date, @date, ?); SET @product_id=LAST_INSERT_ID(); INSERT INTO oc_product_description(product_id, language_id, name, description, tag, meta_title, meta_description, meta_keyword) VALUES(@product_id, ?, ?, ?, ?, ?, ?, ?); INSERT INTO oc_seo_url(store_id, language_id, \`key\`, value, keyword) VALUES(?, ?, ?, @product_id, ?); INSERT INTO oc_product_to_category(product_id, category_id) VALUES(@product_id, ?); INSERT INTO oc_product_to_store(product_id, store_id) VALUES(@product_id, ?); INSERT INTO oc_product_to_layout(product_id, store_id, layout_id) VALUES(@product_id, ?, ?); COMMIT;`
                        
                        await mysql.createConnection({host: '127.0.0.1', user: 'brut', password: 'brut', database: 'ocdb', multipleStatements: true})
                            .then(con => con.query(iSql, [element.master_id, element.model, element.sku, element.upc, element.ean, element.jan, element.isbn, element.mpn, element.location, element.variant, element.override, element.quantity, element.minimum, element.subtract, element.stock_status_id, element.date_available, element.manufacturer_id, element.shipping, element.price, element.points, element.weight, element.weight_class_id, element.length, element.width, element.height, element.length_class_id, element.status, element.tax_class_id, element.sort_order, element.image, element.language_id, element.name, element.description, element.tag, element.meta_title, element.meta_description, element.meta_keyword, element.store_id, element.language_id, element.key, element.keyword, element.category_id, element.store_id, element.store_id, element.layout_id]))
                            .then(ctx.reply('Товары добавлены'))
                            .catch(error => ctx.reply(error));                    
                      };        
                })
            }
        })
    } else {
        await ctx.replyWithHTML('Вы видите эту ошибку, потому что не отправили .xlsx файл или время сессии истекло\n\nПожалуйста, снова отправьте боту .xlsx файл и сразу после загрузки отправьте команду /prodImport')
    }
    
})

bot.on('document', async ctx => {
    const fileId = ctx.message.document.file_id
    const title = ctx.message.document.file_name
    ctx.session.number = ctx.message.document.file_name
    if (ctx.message.document.mime_type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
            await bot.telegram.getFileLink(fileId).then(url => {
                axios({url, responseType:'stream'}).then(response => {
                    return new Promise((resolve, reject) => {
                        response.data.pipe(fs.createWriteStream(`./static/${title}`))
                    })
                })
            }) 
    } else {
        await ctx.reply('Пожалуйста, отправьте нам файл Excel в формате .xlsx')
    }
})

bot.catch((err, ctx) => {
    console.log(err, ctx)
});

bot.launch({
    drop_pending_updates: true
})
.then(() => console.log('Бот успешно стартовал'))

// Enable graceful stop
process.once('SIGINT', () => bot.stop())
process.once('SIGTERM', () => bot.stop())