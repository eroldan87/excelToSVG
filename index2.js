const ExcelJS = require("exceljs");
const fs = require("fs");
const convert = require("xml-js");

const themes = [
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
];

const borderStyles = {
    "medium":{
        "stroke-width":2
    },
    "thin":{},
    "dotted":{
        "stroke-dasharray":3
    },
    "thick":{
        "stroke-width":4
    },
    "mediumDashed":{
        "stroke-width":2,
        "stroke-dasharray":10
    },
    "dashDot":{
        "stroke-dasharray":'10,3,3,3'
    },
    "hair":{
        "stroke-dasharray":2
    },
    "dashDotDot":{
        "stroke-dasharray":'10,3,3,3,3,3'
    },
    "dashed":{
        "stroke-dasharray":'4 2'
    },
    "mediumDashDotDot":{
        "stroke-dasharray":'10,3,3,3,3,3',
        "stroke-width":2
    },
    "slantDashDot":{
        "stroke-dasharray":'10,3,3,3',
        "stroke-width":2
    },
    "mediumDashDot":{
        "stroke-dasharray":'10,3,3,3',
        "stroke-width":2
    },
    "mediumDashed":{
        "stroke-dasharray":'4 2',
        "stroke-width":2
    },
    "double":{
        "stroke-width":4
    }
}

// Patterns
// `darkGray        <pattern id="Pattern" x="0" y="0" width="4" height="4" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="4" height="4" fill="black"/>
// <circle cx="0" cy="0" r="0.5" fill="white"/>
// <circle cx="0" cy="4" r="0.5" fill="white"/>
// <circle cx="2" cy="2" r="0.5" fill="white"/>
// <circle cx="4" cy="0" r="0.5" fill="white"/>
// <circle cx="4" cy="4" r="0.5" fill="white"/>
// </pattern>
// mediumGray  <pattern id="Pattern" x="0" y="0" width="4" height="4" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="4" height="4" fill="black"/>
// <circle cx="0" cy="0" r="0.5" fill="white"/>
// <circle cx="0" cy="4" r="0.5" fill="white"/>
// <circle cx="2" cy="2" r="0.5" fill="white"/>
// <circle cx="4" cy="0" r="0.5" fill="white"/>
// <circle cx="4" cy="4" r="0.5" fill="white"/>
// </pattern>
// lightGray   <pattern id="Pattern" x="0" y="0" width="6" height="6" patternUnits="userSpaceOnUse">
//   <rect x="0" y="0" width="6" height="6" fill="black"/>
//   <circle cx="0" cy="0" r="1" fill="white"/>
//   <circle cx="0" cy="6" r="1" fill="white"/>
//   <circle cx="3" cy="3" r="1" fill="white"/>
//   <circle cx="6" cy="0" r="1" fill="white"/>
//   <circle cx="6" cy="6" r="1" fill="white"/>
// </pattern>
// gray125 <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
//   <rect x="0" y="0" width="8" height="8" fill="black"/>
//   <circle cx="0" cy="0" r="1" fill="white"/>
//   <circle cx="0" cy="8" r="1" fill="white"/>
//   <circle cx="4" cy="4" r="1" fill="white"/>
//   <circle cx="8" cy="0" r="1" fill="white"/>
//   <circle cx="8" cy="8" r="1" fill="white"/>
// </pattern>
// gray0625    
// <pattern id="Pattern" x="0" y="0" width="10" height="10" patternUnits="userSpaceOnUse">
//   <rect x="0" y="0" width="10" height="10" fill="black"/>
//   <circle cx="0" cy="0" r="1" fill="white"/>
//   <circle cx="0" cy="10" r="1" fill="white"/>
//   <circle cx="5" cy="5" r="1" fill="white"/>
//   <circle cx="10" cy="0" r="1" fill="white"/>
//   <circle cx="10" cy="10" r="1" fill="white"/>
// </pattern>
// </defs>
// darkHorizontal  <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
//   <rect x="0" y="0" width="8" height="4" fill="black"/>
//   <rect x="0" y="4" width="8" height="4" fill="white"/>
// </pattern>
// darkVertical    <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="4" height="8" fill="black"/>
// <rect x="4" y="0" width="4" height="8" fill="white"/>
// </pattern>
// darkDown    <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
//   <rect x="0" y="0" width="8" height="8" fill="black"/>
//   <path d="M0 0L8 8" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
//   <path d="M8 0L16 8" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
//   <path d="M0 8L8 16" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
// </pattern>
// darkUp  <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="black"/>
// <path d="M8 0L0 8" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
// <path d="M8 8L0 16" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
// <path d="M0 0L-8 8" fill="" stroke="white" stroke-width="2" stroke-linecap="square"/>
// </pattern>
// darkGrid    <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="black"/>
// <path d="M4 0L4 8" fill="" stroke="white" stroke-width="3" stroke-linecap="square"/>
// <path d="M0 4L8 4" fill="" stroke="white" stroke-width="3" stroke-linecap="square"/>
// </pattern>
// darkTrellis <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="black"/>
// <path d="M8 0L0 8" fill="" stroke="white" stroke-width="3" stroke-linecap="square"/>
// <path d="M0 0L8 8" fill="" stroke="white" stroke-width="3" stroke-linecap="square"/>
// </pattern>
// lightHorizontal <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="7" fill="white"/>
// <rect x="0" y="7" width="8" height="1" fill="black"/>
// </pattern>
// lightVertical   <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="7" height="8" fill="white"/>
// <rect x="7" y="0" width="1" height="8" fill="black"/>
// </pattern>
// lightDown   <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="white"/>
// <path d="M0 0L8 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M8 0L16 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M0 8L8 16" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// </pattern>
// lightUp <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="white"/>
// <path d="M8 0L0 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M8 8L0 16" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M0 0L-8 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// </pattern>
// lightGrid   <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="white"/>
// <path d="M4 0L4 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M0 4L8 4" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// </pattern>
// lightTrellis    <pattern id="Pattern" x="0" y="0" width="8" height="8" patternUnits="userSpaceOnUse">
// <rect x="0" y="0" width="8" height="8" fill="white"/>
// <path d="M8 0L0 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// <path d="M0 0L8 8" fill="" stroke="black" stroke-width="1" stroke-linecap="square"/>
// </pattern>
// `



const workbook = new ExcelJS.Workbook();
workbook.xlsx.readFile("./assets/aa.xlsx").then((res) => {
  // res.eachSheet(function(worksheet, sheetId) {
  //     console.log(sheetId)
  //   });
  /* sheetIds for this file: 51, 71, 60, 50, 66, 67 */

  // console.log(res.getWorksheet(51))

  let sheet1 = res.getWorksheet(51);

  const heights = sheet1._rows.map((x) => {
    return !x.height ? 15 : x.height;
  });

  const widths = sheet1._columns.map((element) => {
    return element.width * 7.45;
  });

  const image = {
    svg: {
      _attributes: {
        version: 1.1,
        width: 20000,
        height: 20000,
        xmlns: "http://www.w3.org/2000/svg",
      },
      defs: {
        linearGradient: [],
        radialGradient: []
      },
      rect: [],
      path: [],
      text: [],
    },
  };

  let top;
  let bottom;
  let left;
  let right;
  let color = "";
  let x;
  let y;
  let height;
  let width;

  const defaultPadding = 3;

  const rows = sheet1._rows;
  
  rows.forEach((row) => {
    row._cells.forEach((cell) => {
        
      if (cell._value._master != null) return;
      if (cell._mergeCount > 0) {
        top = sheet1._merges[cell._address].model.top - 1;
        bottom = sheet1._merges[cell._address].model.bottom;
        left = sheet1._merges[cell._address].model.left - 1;
        right = sheet1._merges[cell._address].model.right;
      } else {
        top = row._number - 1;
        bottom = top + 1;
        left = cell._column._number - 1;
        right = left + 1;
      }
      x =
        widths.slice(0, left).length == 0
          ? 0
          : widths.slice(0, left).reduce((a, b) => a + b);
      y =
        heights.slice(0, top).length == 0
          ? 0
          : heights.slice(0, top).reduce((a, b) => a + b);
      height = heights.slice(top, bottom).reduce((a, b) => a + b);
      width = widths.slice(left, right).reduce((a, b) => a + b);
      x = Math.round(x)
      y = Math.round(y)
      height = Math.round(height)
      width = Math.round(width)


      // For de filling
      if (cell._value.model.style?.fill.fgColor) {
        color = getColor(cell._value.model.style?.fill.fgColor);
      } else if (cell._value.model.style?.fill?.type == "gradient") {
        if (cell._value.model.style.fill.gradient == "angle" || cell._value.model.style.fill.gradient == null) {
          let i = image.svg.defs.linearGradient.findIndex(
            (x) => x.reference == cell._value.model.style.fill
          );
          if (i == -1) {
            image.svg.defs.linearGradient.push({
              _attributes: {
                id: "linearGradient" + cell._address,
                ...getCoords(cell._value.model.style.fill.degree),
              },
              reference: cell._value.model.style.fill,
              stop: cell._value.model.style.fill.stops.map((stop) => {
                return {
                  _attributes: {
                    "stop-color": getColor(stop.color),
                    offset: stop.position,
                  },
                };
              }),
            });
            color = `url(#linearGradient${cell._address})`;
          } else {
            color = `url(#${image.svg.defs.linearGradient[i]._attributes.id})`;
          }
        } else if (cell._value.model.style.fill.gradient == "path") {
            let i = image.svg.defs.radialGradient.findIndex(
                (x) => x.reference == cell._value.model.style.fill
              );
              if (i == -1) {
                image.svg.defs.radialGradient.push({
                  _attributes: {
                    id: "radialGradient" + cell._address,
                    cx: cell._value.model.style.fill.left,
                    cy: cell._value.model.style.fill.top,
                    r:1
                  },
                  reference: cell._value.model.style.fill,
                  stop: cell._value.model.style.fill.stops.map((stop) => {
                    return {
                      _attributes: {
                        "stop-color": getColor(stop.color),
                        offset: stop.position,
                      },
                    };
                  }),
                });
                color = `url(#radialGradient${cell._address})`;
              } else {
                color = `url(#${image.svg.defs.radialGradient[i]._attributes.id})`;
              }
        } else {
          color = null;
        }
      } else {
        color = null;
      }
      
      if (color) {
        image.svg.rect.push({
          _attributes: {
            x: x,
            y: y,
            height: height,
            width: width,
            fill: color,
          },
        });
      }

      //For the borders

      if (cell._value.model.style?.border) {
        Object.entries(cell._value.model.style.border).forEach(([side, value]) => {
            if (cell._address == "AM46") console.log(side, value)
          let color = getColor(value.color);
          color = color == null ? "black" : color;
          let id =
            value.style + "," + color
          let i = image.svg.path.findIndex((x => x._attributes.id == id));
          if (i == -1) {
            i = image.svg.path.push({
              _attributes: {
                id: id,
                d: '',
                stroke: color,
                fill: 'none',
                ...borderStyles[value.style]
              },
            }) - 1;
          }
          switch (side) {
            case "left":
                image.svg.path[i]._attributes.d += 'M' + x + ' ' + y + 'L' + x + ' ' + (y + height)
              break;
            case "right":
                image.svg.path[i]._attributes.d += 'M' + (x + width) + ' ' + y + 'L' + (x + width) + ' ' + (y + height)
              break;
            case "top":
                image.svg.path[i]._attributes.d += 'M' + x + ' ' + y + 'L' + (x + width) + ' ' + y
              break;
            case "bottom":
                image.svg.path[i]._attributes.d += 'M' + x + ' ' + (y + height) + 'L' + (x + width) + ' ' + (y + height)
                break
            case "diagonal":
                image.svg.path[i]._attributes.d += 'M' + x + ' ' + y + 'L' + (x + width) + ' ' + (y + height)
          }
        });
      }

      // For de text
      value =
        cell._value.model.value != null
          ? cell._value.model.value
          : cell._value.model.result != null
          ? cell._value.model.result
          : null;
      if (value != null) {
        text = {
          _attributes: {
            x: Math.round(x + defaultPadding),
            y: Math.round(y + height - defaultPadding),
            "font-size": cell._value.model.style?.font?.size,
          },
          _text: value,
        };
        let color = getColor(cell._value.model.style?.font?.color)
        if (color != null || color != '#000000') text._attributes['fill'] = color
        if (cell._value.model.style?.alignment) {
          switch (cell._value.model.style?.alignment.horizontal) {
            case "right":
              text._attributes.x = Math.round(x + width - defaultPadding);
              text._attributes["text-anchor"] = "end";
              break;
            case "center":
              text._attributes.x = Math.round(x + width / 2);
              text._attributes["text-anchor"] = "middle";
              break;
          }

          switch (cell._value.model.style?.alignment.vertical) {
            case "top":
              text._attributes.y = Math.round(y + defaultPadding);
              text._attributes["dominant-baseline"] = "hanging";
              break;
            case "middle":
              text._attributes.y = Math.round(y + height / 2);
              text._attributes["dominant-baseline"] = "middle";
              break;
          }
        }
        if (cell._value.model.style?.font?.bold)
          text._attributes["font-weight"] = "bold";
        image.svg.text.push(text);
      }
    });
  });

  image.svg.defs.linearGradient = image.svg.defs.linearGradient.map(
    ({ reference, ...theRest }) => theRest
  );
  image.svg.defs.radialGradient = image.svg.defs.radialGradient.map(
    ({ reference, ...theRest }) => theRest
  );

  fs.writeFileSync("data.svg", convert.js2xml(image, { compact: true }));
});

function getColor(color) {
  if (color?.argb) {
    return "#" + color.argb.slice(2, 8);
  } else if (color?.theme != null) {
    const theme = { ...themes[color.theme] };
    const tint = color.tint;
    const changedColor = CalculateFinalLumValue(tint, theme.l);
    theme.l = changedColor;
    return hslToHex(theme);
  } else {
    return null;
  }
}

function hslToHex(hsl) {
  hsl = hsl;
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
    return hex.length === 1 ? "0" + hex : hex;
  };
  return "#" + toHex(r) + toHex(g) + toHex(b);
}

function CalculateFinalLumValue(tint, lum) {
  if (tint == null) return lum;
  if (lum == 0) return lum + tint * 100;
  return (lum * (1.0 + tint) * 240) / 255;
}

function getCoords(angle) {
  if (angle == null) angle = 0;
  angle = (angle * 2 * Math.PI) / 360.0;
  const coords = { x1: 0.5, y1: 0.5, x2: 0, y2: 0 };
  const cuarto = Math.PI / 4;
  if (angle >= cuarto && angle < 3 * cuarto) {
    coords.y2 = 1;
    coords.x2 = (coords.y2 - coords.y1) / Math.tan(angle) + coords.x1;
  } else if (angle >= 3 * cuarto && angle < 5 * cuarto) {
    coords.x2 = 0;
    coords.y2 = (coords.x2 - coords.x1) / Math.tan(angle) + coords.y1;
  } else if (angle >= 5 * cuarto && angle < 7 * cuarto) {
    coords.y2 = 0;
    coords.x2 = (coords.y2 - coords.y1) / Math.tan(angle) + coords.x1;
  } else {
    coords.x2 = 1;
    coords.y2 = (coords.x2 - coords.x1) / Math.tan(angle) + coords.y1;
  }
  coords.y1 = 1 - coords.y2;
  coords.x1 = 1 - coords.x2;
  return coords;
}
