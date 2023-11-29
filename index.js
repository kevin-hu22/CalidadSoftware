//Variable que almacena la instancia de la clase Excel para manipular datos del archivo.
let excel;
//Array que almacena nombres de empleados sin repetir obtenidos del archivo Excel.
let arraySinRepetidos = [];
//Elemento del DOM que representa el input de tipo archivo para cargar el archivo Excel.
const excelInput = document.getElementById('input-excel');
// Elemento del DOM que representa el contenedor de opciones de visualización.
const boxOptions = document.getElementById("box-options");
// NodeList que representa la lista de opciones del menú.
const listOptions = document.querySelectorAll("ul#list-options li button");
const toastContainer = document.getElementById('toast-container');

/**
 * Manejador de eventos para el cambio en la selección de archivos.
 */
excelInput.addEventListener('change', async function () {
    try {
        // Verificar si se seleccionó un archivo
        if (excelInput.files.length > 0) {
            // Obtener la extensión del archivo seleccionado
            const fileExtension = excelInput.files[0].name.split('.').pop().toLowerCase();

            // Verificar si la extensión del archivo es 'xlsx' (Excel)
            if (fileExtension === 'xlsx') {

                // Leer el contenido del archivo Excel y crear una instancia de la clase Excel
                const content = await readXlsxFile(excelInput.files[0])
                excel = new Excel(content)

                // Realizar operaciones con el archivo Excel
                fechaReporte(excel);
                arraySinRepetidos = BuscarPersonal(excel)
                ExcelPrinter.print('head-links', excel)
                crearObjetosEmpleados(arraySinRepetidos)
                extraerInfo()
                mostrarData()
                mostrarAdicionales()
                
            } else {
                const toast = new bootstrap.Toast(toastContainer, {
                    autohide: true, // Puedes ajustar esto según tus necesidades
                });
                toastContainer.textContent = 'Por favor, seleccione un archivo Excel (.xlsx).';
                toast.show();
                // Limpiar la entrada de archivos para permitir al usuario seleccionar otro archivo
                excelInput.value = null;
            }
        }
    } catch (error) {
        // Manejar cualquier error durante la lectura del archivo
        console.error('Error al procesar el archivo:', error.message);
        // Mostrar un mensaje de error utilizando un Toast de Bootstrap
        const toast = new bootstrap.Toast(toastContainer, {
            autohide: true, // Puedes ajustar esto según tus necesidades
        });
        toastContainer.textContent = 'Error al procesar el archivo, es posible que el formato dentro del archivo no sea el adecuado. Por favor, inténtelo de nuevo.';
        toast.show();

        // Limpiar la entrada de archivos para permitir al usuario seleccionar otro archivo
        excelInput.value = null;
    }
});


//Clase que representa el objeto Excel para manipulación de datos.
class Excel {
    constructor(content) {
        this.content = content;
    }
    //Obtiene el encabezado del archivo Excel.
    header() {
        return this.content[0]
    }
    // Obtiene la colección de filas del archivo Excel.
    rows() {
        return new RowsCollection(this.content.slice(1, this.content.length))
    }
    //Obtiene la colección de columnas del archivo Excel.
    columns() {
        return new ColumnsCollection(this.content.slice(1, this.content.length))
    }

}
//Clase que representa la colección de columnas del archivo Excel.
class ColumnsCollection {
    constructor(columns) {
        this.columns = columns;
    }
    //Obtiene la primera columna del archivo Excel.
    firtsCol() {
        return new Col(this.columns[0])
    }
    //Obtiene una columna específica del archivo Excel.
    getCol(index) {
        return new Col(this.columns[index])
    }
    // Obtiene todas las columnas del archivo Excel.
    getCols() {
        return this.columns
    }

    // Obtiene el conteo de columnas del archivo Excel.
    countCols() {
        return this.columns.length
    }
}

//Clase que representa la colección de filas del archivo Excel.
class RowsCollection {
    constructor(rows) {
        this.rows = rows;
    }
    // Obtiene la primera fila del archivo Excel.
    firstRow() {
        return new Row(this.rows[0])
    }

    //Obtiene una fila específica del archivo Excel.
    getRow(index) {
        return new Row(this.rows[index])
    }
    //Obtiene todas las filas del archivo Excel.
    getRows() {
        return this.rows
    }
    // Obtiene el conteo de filas del archivo Excel.
    countRows() {
        return this.rows.length
    }
}
// Clase que representa una fila del archivo Excel.
class Row {
    //Crea una instancia de la clase Row.
    constructor(row) {
        this.row = row
    }
}
// Clase que representa una columna del archivo Excel.
class Col {
    //Crea una instancia de la clase Col.
    constructor(col) {
        this.col = col
    }
}




// Funciones reutilizables


// Hace una busqueda por los Header del archivo para encontrar
// el indice de la columna que se busca, como lo recibe la variable busqueda
// el condicional valida si se encontro devuelve el indice en caso contrario envia un mensaje por consola
function buscarIndex(busqueda) {
    const indexHead = excel.header();
    const index = indexHead.indexOf(busqueda);
    if (index !== -1) {
        return index
    } else {
        return console.log("No se encuentra la columna " + busqueda);
    }
}

// Valida si hay o no informacion de en LocalStorage para trabajar con ella, en caso que no se crea el objeto respectivo
function validarLocalStorage() {
    const objetosEmpleadosGuardados = JSON.parse(localStorage.getItem('objetosEmpleados'));
    // Verificar si existe información en el localStorage, si está actualizada y si la longitud es igual
    if (objetosEmpleadosGuardados && esArrayActualizado(arraySinRepetidos, objetosEmpleadosGuardados) && objetosEmpleadosGuardados.length === arraySinRepetidos.length) {
        return true;
    } else {
        // Si la información no está actualizada, se crean nuevos objetos de empleados
        crearObjetosEmpleados(arraySinRepetidos);
    }
}

// Verifica si cada nombre del arraySinRepetidos está presente en objetosEmpleadosGuardados
function esArrayActualizado(arraySinRepetidos, objetosEmpleadosGuardados) {
    return arraySinRepetidos.every(nombre => {
        return objetosEmpleadosGuardados.some(objeto => objeto.nombre === nombre)
    })
}
// Funciones para el tiempo

//Convierte un tiempo en formato de cadena (HH:MM:SS) a segundos.
// tiempoString - Cadena que representa el tiempo en formato HH:MM:SS.
// Número total de segundos representados por el tiempo proporcionado.
function convertirTiempoASegundos(tiempoString) {
    // Dividir la cadena de tiempo en partes (horas, minutos, segundos)
    const partesTiempo = tiempoString.split(':');
    // Extraer horas, minutos y segundos; si no se proporciona, se asume 0
    const horas = parseInt(partesTiempo[0], 10) || 0;
    const minutos = parseInt(partesTiempo[1], 10) || 0;
    const segundos = parseInt(partesTiempo[2], 10) || 0;
    // Calcular el tiempo total en segundos
    return horas * 3600 + minutos * 60 + segundos;
}

// recibe un array con el tiempo en segundo
function tiempoPro(arrayTiempo) {
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


// Esta funcion recibe un id como identificador del indice posicional, del empleado a buscar
// luego emepieza a mostrar los diferentes datos necesarios con respecto al id
function mostrarData(id) {
    // Elementos HTML donde se mostrará la información
    const mostraAtencion = document.querySelector("#AtencionCant span")
    const mostraNatural = document.querySelector("#NaturalCant span")
    const mostraJuridica = document.querySelector("#JuridicaCant span")
    const exitoTurnos = document.querySelector("#exitoTurnos strong")
    const sinexitoTurnos = document.querySelector("#sinexitoTurnos strong")
    const atenTurnos = document.querySelector("#atenTurnos strong")
    const abanTurnos = document.querySelector("#abanTurnos strong")
    const Pjuridica = document.querySelector("#Pjuridica strong")
    const Pnatural = document.querySelector("#Pnatural strong")
    const tiempoAten = document.querySelector("#tiempoAten strong")
    const tiempoEsp = document.querySelector("#tiempoEsp strong")


    // Validar la existencia y actualización de la información en el localStorage
    if (validarLocalStorage()) {
        // Obtener el array de objetos de empleados almacenado en localStorage
        const arrayDeObjetos = JSON.parse(localStorage.getItem('objetosEmpleados'));

        // Verificar que el índice esté dentro del rango del array
        if (id >= 0 && id < arrayDeObjetos.length) {
            // Obtener el objeto de empleado correspondiente al índice proporcionado
            const objeto = arrayDeObjetos[id];

            // Mostrar los valores en los elementos HTML
            mostraAtencion.innerHTML = objeto.totalTurnos;
            mostraNatural.innerHTML = objeto.pNatural;
            mostraJuridica.innerHTML = objeto.pJuridica;
            exitoTurnos.innerHTML = objeto.solicitudesExitosas;
            sinexitoTurnos.innerHTML = objeto.solicitudesFallidas;
            atenTurnos.innerHTML = objeto.atendidosExito;
            abanTurnos.innerHTML = objeto.cantAbandonados;
            Pjuridica.innerHTML = objeto.pJuridica;
            Pnatural.innerHTML = objeto.pNatural;
            tiempoAten.innerHTML = objeto.tiempoAtencion;
            tiempoEsp.innerHTML = objeto.tiempoEspera;


        } else {
            console.log("Índice fuera de rango");
        }
    }
}
function mostrarAdicionales() {
    const contenido = excel.rows().getRows();
    const indexSerRut = buscarIndex('(SER)Actualización RUT');
    const indexSerInsR = buscarIndex('(SER)Inscripción RUT');
    const indexSerOC = buscarIndex('(SER)Objeto de Campaña');

    const ServRut = document.querySelector("#ServRut strong");
    const ServIrut = document.querySelector("#ServIrut strong");
    const ServOC = document.querySelector("#ServOC strong");
    var RUT = 0;
    var IRUT = 0;
    var OC = 0;

    contenido.forEach(obj => {
        // Para el campo (SER)Actualización RUT
        if (obj[indexSerRut] != null) {
            RUT += 1;
        }

        // Para el campo (SER)Inscripción RUT
        if (obj[indexSerInsR] != null) {
            IRUT += 1;
        }

        // Para el campo (SER)Objeto de Campaña
        if (obj[indexSerOC] != null && obj[indexSerOC] != 'Ninguno') {
            OC += 1;
        }
    });
    console.log(RUT)
    console.log(IRUT)
    console.log(OC)
    ServRut.innerHTML = RUT;
    ServIrut.innerHTML = IRUT;
    ServOC.innerHTML = OC;
}



// Mostrar informacion de usuario selecionado
function userSelected(id) {
    const user = arraySinRepetidos[id]
    document.getElementById("NameUser").innerText = user
    // usa esta funcion para mostrar los datos del usuario selecionado pasandole el id por parametro
    mostrarData(id)
    // consulta si el contenedor en el html tiene active que hace mostrar este contenedor, en caso que no lo tenga se lo añade
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



// itera por cada objeto buscando al sujeto o usuario del cual ira llenado las variables contadoras
// usando los indices especificos para evitar recorrer tantas veces el objeto
// una vez consiga toda la informacion del usuario, utliza la funcion actualizarInformacionEmpleado() enviando por parametros el identificar (nombre) y los datos a actualizados
function extraerInfo() {
    // con las const index... se usa la funcion buscarIndex() para buscar el indice posicional que tiene dicha columna de la cual se busca la informacion
    // se le envia por parametrp el nombre de la columna que se desea buscar
    const indexUser = buscarIndex('NOMBREUSUARIO')
    const indexEstado = buscarIndex('ESTADO')
    const indexTipPersona = buscarIndex('TIPOCLIENTE')
    const indexResultado = buscarIndex('(SER)Resultado del Tramite')
    const indexTiempoTotal = buscarIndex('TOTAL')
    const indexTiempoAten = buscarIndex('ATENCION')
    const indexTiempoEsp = buscarIndex('ESPERA')
    // extrae todas las filas del doc y los convierte en un array de objetos
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

    // se itera por cada nombre de empleados previamente filtrados en el array "arraySinRepetidos"
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

        // esta funcion recorre todos los objetos en busca del cual coincida con el nombre por el cual se itera
        informacion.forEach(obj => {
            const nombre = obj[indexUser];
            if (nombre === nombreEnviar) {
                turnos += 1;
                //esta condicion consulta si el valor que encontro es igual a "Finalizado" o no para añadir a las variables contadoras
                if (obj[indexEstado] === "Finalizado") {
                    cantAtendido += 1;
                } else {
                    cantAban += 1;
                }
                //esta condicion consulta si el valor que encontro es igual a "Persona Natural" o no para añadir a las variables contadoras
                if (obj[indexTipPersona] === "Persona Natural") {
                    pNatural += 1;
                } else {
                    pJuridica += 1;
                }
                //esta condicion consulta si el valor que encontro es igual a "Exitoso" o no para añadir a las variables contadoras
                if (obj[indexResultado] === "Exitoso") {
                    cantExi += 1;
                } else {
                    cantFall += 1;
                }
                totalTiempo.push(obj[indexTiempoTotal])
                // esta condicion eviata que al array de tiempo de Atencion ingresen valores iguales a "00:00:00", para evitar errores futuros
                if (obj[indexTiempoAten] !== "00:00:00") {
                    tiempoAtend.push(obj[indexTiempoAten])
                }
                tiempoEsp.push(obj[indexTiempoEsp])
            }
        })
        // luego de llenar las variables, se ejecuta la funcion para actualizar la informacion en el localStorage
        actualizarInformacionEmpleado(nombreEnviar, turnos, cantAtendido, pNatural, pJuridica, cantAban, cantExi, cantFall, totalTiempo, tiempoAtend, tiempoEsp)
    }
}

// recibe la informacion por parametros para actualizar la informacion en el Array de Objetos guardado en LocalStorage
// la funcion cuenta con los siguientes parametros, (nombre) este es el identificador para buscar el objeto con este nombre especifico
// por consiguiente recibe todos los datos que se van a actualizar
function actualizarInformacionEmpleado(nombre, turnos, cantAtendido, pNatural, pJuridica, cantAban, cantExi, cantFall, totalTiempo, tiempoAtend, tiempoEsp) {
    const objetosEmpleadosGuardados = JSON.parse(localStorage.getItem('objetosEmpleados'));
    const empleado = objetosEmpleadosGuardados.find(e => e.nombre === nombre);
    let tiempoMAten = [];

    const tiempoMTotal = totalTiempo.map(tiempo => convertirTiempoASegundos(tiempo))
    // Consulta si el array de Atencion llego vacio, de ser asi el array le asigna 0
    if (tiempoAtend.length === 0) {
        tiempoMAten = [0];
    } else {
        tiempoMAten = tiempoAtend.map(tiempo => convertirTiempoASegundos(tiempo))
    }
    const tiempoMEsp = tiempoEsp.map(tiempo => convertirTiempoASegundos(tiempo))
    // calcular promdio de tiempos
    const proTotal = tiempoPro(tiempoMTotal)
    const proAten = tiempoPro(tiempoMAten)
    const proEsp = tiempoPro(tiempoMEsp)

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

// muestra la informacion en este caso crear los botones con los nombres de cada usuario
class ExcelPrinter {
    static print(headLinksId, excel) {
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
    }
}

/**
 * Función que busca y retorna los nombres de usuarios en el archivo Excel.
 * @param {Excel} excel - Instancia de la clase Excel que representa el contenido del archivo.
 * @returns {Array<string>} - Array con los nombres de usuarios encontrados y sin duplicados.
 */
function BuscarPersonal(excel) {
    // Obtener todas las filas del archivo Excel
    const TotalRegistros = excel.rows().getRows();

    // Array para almacenar los nombres de usuarios
    const nombresUsuarios = [];

    // Obtener el índice de la columna 'NOMBREUSUARIO'
    const userIndex = buscarIndex('NOMBREUSUARIO');

    // Iterar sobre cada fila del archivo
    TotalRegistros.forEach(fila => {
        // Obtener el nombre de usuario de la fila
        const nombreUsuario = fila[userIndex];

        // Verificar si el nombre de usuario existe y no es "Sin Usuario"
        if (nombreUsuario && nombreUsuario !== "Sin Usuario") {
            // Agregar el nombre de usuario al array
            nombresUsuarios.push(nombreUsuario);
        }
    });

    // Eliminar elementos duplicados en el array de nombres de usuarios
    const nombresUnicos = Array.from(new Set(nombresUsuarios));
    return nombresUnicos;
}
