let excel;
let arraySinRepetidos = [];
const excelInput = document.getElementById('input-excel');
const boxOptions = document.getElementById("box-options");
const listOptions = document.querySelectorAll("ul#list-options li.op button");

// Lectura del archivo
excelInput.addEventListener('change', async function () {
    const content = await readXlsxFile(excelInput.files[0])
    excel = new Excel(content)

    fechaReporte(excel);
    arraySinRepetidos = BuscarPersonal(excel)
    ExcelPrinter.print('tablaInfo', 'head-links', excel)

    crearObjetosEmpleados(arraySinRepetidos)
    extraerInfo()
    mostrarData()

})


class Excel {
    constructor(content) {
        this.content = content;
    }
    header() {
        return this.content[0]
    }
    rows() {
        return new RowsCollection(this.content.slice(1, this.content.length))
    }
    columns() {
        return new ColumnsCollection(this.content.slice(1, this.content.length))
    }

}
class ColumnsCollection {
    constructor(columns) {
        this.columns = columns;
    }
    firtsCol() {
        return new Col(this.columns[0])
    }
    //trae todos los datos
    getCol(index) {
        return new Col(this.columns[index])
    }
    getCols() {
        return this.columns
    }

    // trae el conteo de valores
    countCols() {
        return this.columns.length
    }
}


class RowsCollection {
    constructor(rows) {
        this.rows = rows;
    }
    // trae el primer dato
    firstRow() {
        return new Row(this.rows[0])
    }

    //trae todos los datos
    getRow(index) {
        return new Row(this.rows[index])
    }
    getRows() {
        return this.rows
    }
    // trae el conteo de valores
    countRows() {
        return this.rows.length
    }
}
class Row {
    constructor(row) {
        this.row = row
    }

    asesorNombre(excel) {
        const header = excel.header();
        const index = header.indexOf('NOMBREUSUARIO');

        if (index !== -1) {
            return this.row[index];
        } else {
            return console.log("No se encuentra la columna de Nombre Usuario");
        }
    }
    tiempoTotal() {
        return this.row[28];
    }
    estado() {
        return this.row[32];
    }
    tipoPersona() {
        return this.row[8];
    }
}
class Col {
    constructor(col) {
        this.col = col
    }

    asesorNombre(excel) {
        const header = excel.header();
        const index = header.indexOf('NOMBREUSUARIO');

        if (index !== -1) {
            return this.col[index];
        } else {
            return console.log("No se encuentra la columna 'NOMBREUSUARIO'");
        }
    }
    tiempoTotal(excel) {
        const header = excel.header();
        const index = header.indexOf('TOTAL');

        if (index !== -1) {
            return this.col[index];
        } else {
            return console.log("No se encuentra la columna 'TOTAL'");
        }
    }
    estado(excel) {
        const header = excel.header();
        const index = header.indexOf('ESTADO');

        if (index !== -1) {
            return this.col[index];
        } else {
            return console.log("No se encuentra la columna 'ESTADO'");
        }
    }
    tipoPersona(excel) {
        const header = excel.header();
        const index = header.indexOf('TIPOCLIENTE');

        if (index !== -1) {
            return this.col[index];
        } else {
            return console.log("No se encuentra la columna 'TIPOCLIENTE'");
        }
    }
    aniones(excel) {
        const header = excel.header()
        const index = header.indexOf('ANIOMES')

        if (index !== -1) {
            return this.col[index];
        } else {
            return console.log("No se encuentra la columna 'TIPOCLIENTE'");
        }
    }
}




// Funciones reutilizables
function buscarIndex(busqueda) {
    const indexHead = excel.header();
    const index = indexHead.indexOf(busqueda);
    if (index !== -1) {
        return index
    } else {
        return console.log("No se encuentra la columna " + busqueda);
    }
    // Hace una busqueda por los Header del archivo para encontrar
    // el indice de la columna que se busca, como lo recibe la variable busqueda
    // el condicional valida si se encontro devuelve el indice en caso contrario envia un mensaje por consola
}

function validarLocalStorage(){
    const objetosEmpleadosGuardados = JSON.parse(localStorage.getItem('objetosEmpleados'))

    if (objetosEmpleadosGuardados && esArrayActualizado(arraySinRepetidos, objetosEmpleadosGuardados) && objetosEmpleadosGuardados.length == arraySinRepetidos.length) {
        return true
    }else{
        crearObjetosEmpleados(arraySinRepetidos)
    }
}
function mostrarData(id) {
    const mostraAtencion = document.querySelector("#AtencionCant span");
    const mostraNatural = document.querySelector("#NaturalCant span");
    const mostraJuridica = document.querySelector("#JuridicaCant span");

    if (validarLocalStorage()) {
        const arrayDeObjetos = JSON.parse(localStorage.getItem('objetosEmpleados'));

        // Verificar que el índice esté dentro del rango del array
        if (id >= 0 && id < arrayDeObjetos.length) {
            const objeto = arrayDeObjetos[id];

            // Mostrar los valores en los elementos HTML
            mostraAtencion.innerHTML = objeto.totalTurnos;
            mostraNatural.innerHTML = objeto.pNatural;
            mostraJuridica.innerHTML = objeto.pJuridica;
        } else {
            console.log("Índice fuera de rango");
        }
    }
}


// Mostrar informacion de usuario selecionado
function userSelected(id) {
    const user = arraySinRepetidos[id]
    document.getElementById("NameUser").innerText = user
    
    mostrarData(id)

    if (!boxOptions.classList.contains('active')) {
        boxOptions.classList.add('active')
    }

}

// Recuperar el array de objetos del localStorage
// Verificar si el array de objetos existe y está actualizado
// Si está actualizado, retornar el array guardado
// Si no existe o no está actualizado, crear el nuevo array de objetos
// Crea un objeto con la estructura específica para cada nombre de empleado
// Agrega el objeto creado al array de objetos
// Almacena el nuevo array de objetos en el localStorage
function crearObjetosEmpleados(arraySinRepetidos) {
    const objetosEmpleadosGuardados = JSON.parse(localStorage.getItem('objetosEmpleados'))

    if (objetosEmpleadosGuardados && esArrayActualizado(arraySinRepetidos, objetosEmpleadosGuardados) && objetosEmpleadosGuardados.length == arraySinRepetidos.length) {
        return objetosEmpleadosGuardados;
    }
    const objetosEmpleados = [];
    arraySinRepetidos.forEach(nombre => {
        const empleado = {
            nombre: nombre,
            totalTurnos: 0,          // turnos al presente dia para este usuario
            atendidosExito: 0,      // procesos de atencion finalizados
            pNatural: 0,            // atencion de persona naturales    
            pJuridica: 0,           // atencion de persona juridicas
            cantAbandonados: 0,     // procesos de atencion abandonados
            solicitudesExitosas: 0, // Procesos de atencion con exito en el tramite
            solicitudesFallidas: 0, // Procesos de atencion con fracaso para finalizar por falta de docuemntos, etc
            tiempoTotal: 0,         // Tiempo de atencion (suma de todos los valores)
            tiempoAtencion: 0,      // Tiempo de atencion (suma de todos los valores) 
            tiempoEspera: 0,        // Tiempo de espera (suma de todos los valores)
        }
        objetosEmpleados.push(empleado);
    })
    localStorage.setItem('objetosEmpleados', JSON.stringify(objetosEmpleados));
    return objetosEmpleados;
}

// Verifica si cada nombre en arraySinRepetidos está presente en objetosEmpleadosGuardados
function esArrayActualizado(arraySinRepetidos, objetosEmpleadosGuardados) {
    return arraySinRepetidos.every(nombre => {
        return objetosEmpleadosGuardados.some(objeto => objeto.nombre === nombre)
    })
}


// itera por cada objeto buscando al sujeto o usuario del cual ira llenado las variables contadoras
// usando los indices especificos para evitar recorrer tantas veces el objeto
// una vez consiga toda la informacion del usuario, utliza la funcion actualizarInformacionEmpleado() enviando por parametros el identificar (nombre) y los datos a actualizados
function extraerInfo() {
    const indexUser = buscarIndex('NOMBREUSUARIO')
    const indexEstado = buscarIndex('ESTADO')
    const indexTipPersona = buscarIndex('TIPOCLIENTE')
    const indexResultado = buscarIndex('(SER)Resultado del Tramite')
    const indexTiempoTotal = buscarIndex('TOTAL')
    const indexTiempoAten = buscarIndex('ATENCION')
    const indexTiempoEsp = buscarIndex('ESPERA')
    const informacion = excel.rows().getRows();

    let nombreEnviar = '';
    let turnos = 0;
    let cantAtendido = 0;
    let pNatural = 0;
    let pJuridica = 0;
    let cantAban = 0;
    let cantExi = 0;
    let cantFall = 0;
    let totalTiempo = [];
    let tiempoAtend = [];
    let tiempoEsp = [];


    for (let i = 0; i < arraySinRepetidos.length; i++) {
        turnos = 0;
        cantAtendido = 0;
        pNatural = 0;
        pJuridica = 0;
        cantAban = 0;
        cantExi = 0;
        cantFall = 0;
        totalTiempo = [];
        tiempoAtend = [];
        tiempoEsp = [];
        nombreEnviar = arraySinRepetidos[i];

        informacion.forEach(obj => {
            const nombre = obj[indexUser];
            if (nombre === nombreEnviar) {
                turnos += 1;
                if (obj[indexEstado] === "Finalizado") {
                    cantAtendido += 1;
                } else {
                    cantAban += 1;
                }
                if (obj[indexTipPersona] === "Persona Natural") {
                    pNatural += 1;
                } else {
                    pJuridica += 1;
                }
                if (obj[indexResultado] === "Exitoso") {
                    cantExi += 1;
                } else {
                    cantFall += 1;
                }
                totalTiempo.push(obj[indexTiempoTotal])

                if (obj[indexTiempoAten] !== "00:00:00") {
                    tiempoAtend.push(obj[indexTiempoAten])
                }
                tiempoEsp.push(obj[indexTiempoEsp])
            }
        })
        
        actualizarInformacionEmpleado(nombreEnviar, turnos, cantAtendido, pNatural, pJuridica, cantAban, cantExi, cantFall, totalTiempo, tiempoAtend, tiempoEsp)
    }
}

function actualizarInformacionEmpleado(nombre, turnos, cantAtendido, pNatural, pJuridica, cantAban, cantExi, cantFall, totalTiempo, tiempoAtend, tiempoEsp) {
    const objetosEmpleadosGuardados = JSON.parse(localStorage.getItem('objetosEmpleados'));
    const empleado = objetosEmpleadosGuardados.find(e => e.nombre === nombre);
    let tiempoMAten = [];

    const tiempoMTotal = totalTiempo.map(tiempo => convertirTiempoASegundos(tiempo))
    if(tiempoAtend.length === 0){
         tiempoMAten = [0];
    }else{
        tiempoMAten = tiempoAtend.map(tiempo => convertirTiempoASegundos(tiempo))
    }
    const tiempoMEsp = tiempoEsp.map(tiempo => convertirTiempoASegundos(tiempo))
    // calcular promdio de tiempos
    const proTotal = tiempoPro(tiempoMTotal)
    const proAten = tiempoPro(tiempoMAten)
    const proEsp = tiempoPro(tiempoMEsp)
    /* console.log("promedio de tiempo: ")
    console.log(`El tiempo promedio de Total fue de ${proTotal} minutos.`);
    console.log(`El tiempo promedio de atención fue de ${proAten} minutos.`);
    console.log(`El tiempo promedio de Espera fue de ${proEsp} minutos.`); */

    if (empleado) {
        // Actualiza la información del empleado si se encuentra en el array
        empleado.totalTurnos = turnos;
        empleado.atendidosExito = cantAtendido;
        empleado.pNatural = pNatural;
        empleado.pJuridica = pJuridica;
        empleado.cantAbandonados = cantAban;
        empleado.solicitudesExitosas = cantExi;
        empleado.solicitudesFallidas = cantFall;
        empleado.tiempoTotal = proTotal;
        empleado.tiempoAtencion = proAten;
        empleado.tiempoEspera = proEsp;

        localStorage.setItem('objetosEmpleados', JSON.stringify(objetosEmpleadosGuardados));
    } else {
        // Si el empleado no se encuentra, puedes manejarlo de acuerdo a tus necesidades
        console.log(`No se encontró al empleado con el nombre ${nombre}`);
    }
}

// Inicializar las variables para horas, minutos y segundos totales
// Iterar sobre cada tiempo en el array      
function sumarTiemposEnArray(arrayTiempos) {
    let totalHoras = 0;
    let totalMinutos = 0;
    let totalSegundos = 0;

    arrayTiempos.forEach(tiempoString => {
        const partesTiempo = tiempoString.split(':'); // Dividir el string en partes (horas, minutos, segundos)

        // Convertir cada parte a un entero
        const horas = parseInt(partesTiempo[0], 10) || 0; // Si no se puede convertir, asume 0
        const minutos = parseInt(partesTiempo[1], 10) || 0;
        const segundos = parseInt(partesTiempo[2], 10) || 0;

        // Sumar las horas, minutos y segundos al total
        totalHoras += horas;
        totalMinutos += minutos;
        totalSegundos += segundos;
    });

    // Convertir exceso de minutos y segundos
    totalMinutos += Math.floor(totalSegundos / 60);
    totalSegundos %= 60;

    totalHoras += Math.floor(totalMinutos / 60);
    totalMinutos %= 60;

    // Devolver el tiempo total en formato HH:MM:SS
    const tiempoTotal = `${formatoDosDigitos(totalHoras)}:${formatoDosDigitos(totalMinutos)}:${formatoDosDigitos(totalSegundos)}`;

    return tiempoTotal;
}
function convertirTiempoASegundos(tiempoString) {
    const partesTiempo = tiempoString.split(':');
    const horas = parseInt(partesTiempo[0], 10) || 0;
    const minutos = parseInt(partesTiempo[1], 10) || 0;
    const segundos = parseInt(partesTiempo[2], 10) || 0;
    return horas * 3600 + minutos * 60 + segundos;
}
function tiempoPro(arrayTiempo){
    const tiempoRes = arrayTiempo.map(tiempoSegundos => tiempoSegundos / 60);
    const tiempoPromedio = tiempoRes.reduce((total, tiempo) => total + tiempo, 0) / tiempoRes.length;
    const resultado = tiempoPromedio.toFixed(2);
    
    return resultado

}


// Obtiene la información de la columna correspondiente a 'FECHACREACION'.
// Extrae la parte de la fecha hasta el primer espacio.
// Actualiza el elemento HTML con la fecha formateada.
function fechaReporte(excel) {
    const indexFecha = buscarIndex('FECHACREACION')
    const fechaCompleta = excel.columns().getCol(0).col[indexFecha].toString();
    const formateada = fechaCompleta.slice(0, fechaCompleta.indexOf(' '));
    document.querySelector("h4.fecha span").innerText = formateada;
}

// muestra la informacion en este caso crear los botones con los nombres de cada usuario
class ExcelPrinter {
    static print(tableId, headLinksId, excel) {
        const table = document.getElementById(tableId)
        const headLinks = document.getElementById(headLinksId)
   
        const personal = arraySinRepetidos;

        personal.forEach((User, index) => {
            headLinks.querySelector("div").innerHTML += `<button class="btn" id="${index}" onclick="userSelected(this.id)"><svg width="32" height="32" viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg">
            <g id="iconamoon:profile">
            <g id="Group">
            <path id="Vector" d="M5.33301 23.9993C5.33301 22.5849 5.89491 21.2283 6.8951 20.2281C7.8953 19.2279 9.25185 18.666 10.6663 18.666H21.333C22.7475 18.666 24.104 19.2279 25.1042 20.2281C26.1044 21.2283 26.6663 22.5849 26.6663 23.9993C26.6663 24.7066 26.3854 25.3849 25.8853 25.885C25.3852 26.3851 24.7069 26.666 23.9997 26.666H7.99967C7.29243 26.666 6.61415 26.3851 6.11406 25.885C5.61396 25.3849 5.33301 24.7066 5.33301 23.9993Z" stroke="#589D8F" stroke-width="2" stroke-linejoin="round"/>
            <path id="Vector_2" d="M16 13.333C18.2091 13.333 20 11.5421 20 9.33301C20 7.12387 18.2091 5.33301 16 5.33301C13.7909 5.33301 12 7.12387 12 9.33301C12 11.5421 13.7909 13.333 16 13.333Z" stroke="#589D8F" stroke-width="2"/>
            </g>
            </g>
            </svg>${User}</button>`
        });

        /* 
        const array = [1, 4, 5, 100]

                for (const index of array) {
                    const row = excel.columns().getCol(index);
        
                    table.querySelector('tbody').innerHTML += `
                <tr> 
                    <td>${row.asesorNombre(excel)}</td>
                    
                </tr>
            `;
                } */
        /* 

                for (let index = 0; index < excel.rows().count(); index++) {
                    const row = excel.rows().get(index);
        
                    table.querySelector('tbody').innerHTML += `
                    <tr> 
                    <td>${row.asesorNombre()}</td>
                    <td>${row.tiempoTotal()}</td>            
                    <td>${row.estado()}</td>
                    <td>${row.tipoPersona()}</td>
                    </tr>`
                } */
    }
}


function CargarInfo(id) {
    EvaluarTiempo(id)
}

function EvaluarTiempo() {
    const listaPersonal = arraySinRepetidos;
    console.log("Valores no repetidos: " + listaPersonal)


}

function BuscarPersonal(excel) {
    const TotalRegistros = excel.rows().getRows();
    const nombresUsuarios = [];
    const userIndex = buscarIndex('NOMBREUSUARIO');

    TotalRegistros.forEach(fila => {
        const nombreUsuario = fila[userIndex];
        if (nombreUsuario && nombreUsuario !== "Sin Usuario") {
            nombresUsuarios.push(nombreUsuario);

        }
    });

    // Eliminar elementos duplicados
    const arrayListo = Array.from(new Set(nombresUsuarios));
    return arrayListo;
}



/*  console.log('Asesor: ' + excel.rows().first().asesorNombre())
 console.log('Tiempo: ' + excel.rows().first().tiempoTotal())
 console.log('Estado: ' + excel.rows().first().estado())
 console.log('Tipo persona: ' + excel.rows().first().tipoPersona()) */