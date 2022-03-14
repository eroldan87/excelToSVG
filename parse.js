
// const convert = require('xml-js');

// const xml = `<cjuego><juego>Bingo Story</juego><solicitudJuegosOS><tipoMaquina>BLUE</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego>
// <cjuego><juego>La Calaca Bingo</juego><solicitudJuegosOS><tipoMaquina>WBR2</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>Fairy Fantasy</juego><solicitudJuegosOS><tipoMaquina>SLEIC</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>DISCOBALL</juego><solicitudJuegosOS><tipoMaquina>BLUE</tipoMaquina><cantidad>2</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>Fire Bingo</juego><solicitudJuegosOS><tipoMaquina>WBR1</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>DISCOBALL</juego><solicitudJuegosOS><tipoMaquina>BLUE</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>Fire Bingo</juego><solicitudJuegosOS><tipoMaquina>WBR1</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego><cjuego><juego>Bingo 
// Story</juego><solicitudJuegosOS><tipoMaquina>BLUE</tipoMaquina><cantidad>1</cantidad></solicitudJuegosOS></cjuego>`

// console.log(convert.xml2js(null, {compact:true, ignoreDeclaration:true}))

function getCoords(angle) {
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

console.log({id:'hola',...getCoords(45)})