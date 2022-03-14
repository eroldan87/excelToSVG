const ExcelJS = require('exceljs');
const fs = require('fs')
const CircularJSON = require('circular-json')
const convert = require('xml-js');

themes = [
    { h: 0, s: 0, l: 100 },
    { h: 0, s: 0, l: 0 },
    { h: 51, s: 28, l: 91 },
    { h: 231, s: 60, l: 31 },
    { h: 213, s: 45, l: 53 },
    { h: 2, s: 48, l: 53 },
    { h: 80, s: 42, l: 54 },
    { h: 267, s: 25, l: 51 },
    { h: 193, s: 52, l: 54 },
    { h: 27, s: 92, l: 62 },
]

// var workbook = XLSX.readFile('./assets/aa.xlsx',{cellStyles:true});

// console.log(workbook.Sheets[workbook.SheetNames[0]])
// fs.writeFileSync('data.json', JSON.stringify(workbook.Sheets[workbook.SheetNames[0]]))
// const html = XLSX.utils.sheet_to_html(workbook.Sheets[workbook.SheetNames[0]])
// fs.writeFileSync('data.html', html)

const workbook = new ExcelJS.Workbook()
workbook.xlsx.readFile('./assets/Libro2.xlsx').then(res => {
    // res.eachSheet(function(worksheet, sheetId) {
    //     console.log(sheetId)
    //   });
    /* sheetIds for this file: 51, 71, 60, 50, 66, 67 */

    // console.log(res.getWorksheet(51))

    let sheet1 = res.getWorksheet(1)
    delete sheet1._workbook

    const heights = sheet1._rows.map(x => {
        return !x.height ? 15 : x.height
    })

    const widths = sheet1._columns.map(element => {
        return element.width * 7.45
    });

    let image = {
        svg: {
            _attributes: {
                version: 1.1,
                width: 20000,
                height: 20000,
                xmlns: "http://www.w3.org/2000/svg"
            },
            defs: {
                linearGradient: []
            },
            rect: []
        }
    }

    let rows = sheet1._rows.map(x => {
        return x._cells.filter(y => {
            return !y._value._master
        })
            .map(y => {
                return { row: x._number, column: y._column._number, value: y._value, address: y._address }
            })
    }).flat()

    Object.keys(sheet1._merges).forEach(key => {
        let top = sheet1._merges[key].model.top - 1
        let bottom = sheet1._merges[key].model.bottom
        let left = sheet1._merges[key].model.left - 1
        let right = sheet1._merges[key].model.right
        let cell = rows.find(x => {
            return x.address == key
        })
        cell['checked'] = true
        // console.log(cell.value.model.style)

        let color = ''
        if (cell.value.model.style?.fill?.fgColor?.argb) {
            color = '#' + cell.value.model.style.fill.fgColor.argb.slice(2, 8)
        } else if (cell.value.model.style?.fill?.fgColor?.theme) {
            let theme = { ...themes[cell.value.model.style.fill.fgColor.theme] }
            let tint = cell.value.model.style.fill.fgColor.tint
            let changedColor = CalculateFinalLumValue(tint, theme.l)
            theme.l = changedColor
            color = hslToHex(theme)
        } else {
            color = 'white'
        }

        image.svg.rect.push(
            {
                _attributes: {
                    x: widths.slice(0, left).length == 0 ? 0 : widths.slice(0, left).reduce((a, b) => a + b),
                    y: heights.slice(0, top).length == 0 ? 0 : heights.slice(0, top).reduce((a, b) => a + b),
                    height: heights.slice(top, bottom).reduce((a, b) => a + b),
                    width: widths.slice(left, right).reduce((a, b) => a + b),
                    fill: color,
                    // stroke: "black",
                    // "stroke-width": 3
                }
            }
        )
        // console.log(element.model)
    });

    // console.log(rows.find(x => x.value.model.style == rows.find(x => x.address == 'P25').value.model.style))
    // console.log(rows.find(x => x.address == 'BX28').value.model.style)

    rows.forEach(cell => {
        if (cell.checked) return
        let top = cell.row - 1
        let left = cell.column - 1
        if (cell.address == 'O28') console.log(JSON.stringify(cell.value.model.style.fill))

        let color = ''
        if (cell.value.model.style?.fill.fgColor) {
            color = getColor(cell.value.model.style?.fill.fgColor)
        } else if (cell.value.model.style?.fill?.type == 'gradient') {
            if (cell.value.model.style.fill.gradient == 'angle') {
                let i = image.svg.defs.linearGradient.findIndex(x => x.reference == cell.value.model.style.fill)
                if (i == -1) {
                    image.svg.defs.linearGradient.push({
                        _attributes: {
                            id: 'linearGradient' + cell.address,
                            ...getCoords(cell.value.model.style.fill.degree)
                        },
                        reference: cell.value.model.style.fill,
                        stop: cell.value.model.style.fill.stops.map(stop => {
                            return {
                                _attributes: {
                                    "stop-color": getColor(stop.color),
                                    offset:stop.position
                                }
                            }
                        })
                    })
                    color = `url(#linearGradient${cell.address})`
                } else {
                    color = `url(#${image.svg.defs.linearGradient[i]._attributes.id})`
                }
            } else {
                color = 'black'
            }
        } else {
            return
        }

        image.svg.rect.push(
            {
                _attributes: {
                    x: widths.slice(0, left).length == 0 ? 0 : widths.slice(0, left).reduce((a, b) => a + b),
                    y: heights.slice(0, top).length == 0 ? 0 : heights.slice(0, top).reduce((a, b) => a + b),
                    height: heights[top],
                    width: widths[left],
                    fill: color
                }
            }
        )

    })

    //     console.log(JSON.stringify(convert.xml2js(svg, {compact:true})))




    fs.writeFileSync('data.svg', convert.js2xml(image, { compact: true }))
    fs.writeFileSync('data.json', CircularJSON.stringify(rows))

})
// fs.writeFileSync('data.json', JSON.stringify(Workbook))
function getColor(color) {
    if (color.argb) return '#' + color.argb.slice(2, 8)
    if (color.theme != null) {
        const theme = { ...themes[color.theme] }
        const tint = color.tint
        const changedColor = CalculateFinalLumValue(tint, theme.l)
        theme.l = changedColor
        return hslToHex(theme)

    } else if (color.indexed != null) {
        return 'black'
    } 
}

function hslToHex(hsl) {
    hsl = hsl
    let h = hsl.h;
    let s = hsl.s;
    let l = hsl.l;
    h /= 360;
    s /= 100;
    l /= 100;
    let r, g, b;
    if (s === 0) {
        r = g = b = l;
    } else {
        const hue2rgb = function (p, q, t) {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1 / 6) return p + (q - p) * 6 * t;
            if (t < 1 / 2) return q;
            if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
            return p;
        };
        const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        const p = 2 * l - q;
        r = hue2rgb(p, q, h + 1 / 3);
        g = hue2rgb(p, q, h);
        b = hue2rgb(p, q, h - 1 / 3);
    }
    const toHex = function (x) {
        const hex = Math.round(x * 255).toString(16);
        return hex.length === 1 ? '0' + hex : hex;
    };
    return '#' + toHex(r) + toHex(g) + toHex(b);
}

function CalculateFinalLumValue(tint, lum) {

    if (tint == null) return lum;
    if (lum == 0) return lum + tint * 100
    return lum * (1.0 + tint);

} 

function getCoords(angle) {
    if (angle == null) angle = 0
    angle = angle * 2 * Math.PI / 360.0
    const coords = {x1: 0.5, y1: 0.5, x2:0, y2:0}
    const cuarto = Math.PI / 4.
    if (angle >= cuarto && angle < 3 * cuarto) {
        coords.y2 = 1
        coords.x2 = (coords.y2 - coords.y1) / Math.tan(angle) + coords.x1
    } else if (angle >= 3 * cuarto && angle < 5 * cuarto) {
        coords.x2 = 0
        coords.y2 = (coords.x2 - coords.x1) / Math.tan(angle) + coords.y1
    } else if (angle >= 5 * cuarto && angle < 7 * cuarto) {
        coords.y2 = 0
        coords.x2 = (coords.y2 - coords.y1) / Math.tan(angle) + coords.x1
    } else {
        coords.x2 = 1
        coords.y2 = (coords.x2 - coords.x1) / Math.tan(angle) + coords.y1
    }
    coords.y1 = 1-coords.y2
    coords.x1 = 1-coords.x2
    return coords
}