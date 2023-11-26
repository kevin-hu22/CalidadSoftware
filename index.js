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

    // Llama a la función y obtén el array de objetos
    const empleadosConInfo = crearObjetosEmpleados(arraySinRepetidos);

    // Muestra el resultado
    console.log(empleadosConInfo);

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


function crearObjetosEmpleados(arraySinRepetidos) {
    // Recuperar el array de objetos del localStorage
    const objetosEmpleadosGuardados = JSON.parse(localStorage.getItem('objetosEmpleados'));

    // Verificar si el array de objetos existe y está actualizado
    if (objetosEmpleadosGuardados && esArrayActualizado(arraySinRepetidos, objetosEmpleadosGuardados) && objetosEmpleadosGuardados.length == arraySinRepetidos.length) {
        // Si está actualizado, retornar el array guardado
        return objetosEmpleadosGuardados;
    }

    // Si no existe o no está actualizado, crear el nuevo array de objetos
    const objetosEmpleados = [];

    arraySinRepetidos.forEach(nombre => {
        // Crea un objeto con la estructura específica para cada nombre de empleado
        const empleado = {
            nombre: nombre,
            atendidosExito: 0,
            pNatural: 0,
            pJuridica: 0,
            cantAbandonados: 0,
            solicitudesExitosas: 0,
            solicitudesFallidas: 0,
            tiempoTotal: 0,
            tiempoAtencion: 0
        };

        // Agrega el objeto creado al array de objetos
        objetosEmpleados.push(empleado);
    });

    // Almacena el nuevo array de objetos en el localStorage
    localStorage.setItem('objetosEmpleados', JSON.stringify(objetosEmpleados));

    return objetosEmpleados;
}

function esArrayActualizado(arraySinRepetidos, objetosEmpleadosGuardados) {
    // Verifica si cada nombre en arraySinRepetidos está presente en objetosEmpleadosGuardados
    return arraySinRepetidos.every(nombre => {
        return objetosEmpleadosGuardados.some(objeto => objeto.nombre === nombre);
    });
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



// Mostrar informacion de usuario selecionado
function userSelected(id) {
    const user = arraySinRepetidos[id]
    document.getElementById("NameUser").innerText = user
    



    if (!boxOptions.classList.contains('active')) {
        boxOptions.classList.add('active')
    }

}

// Retorna la cantidad de personas atendidas con exito
function atentidosHoy(user) {
    const indexUser = buscarIndex('NOMBREUSUARIO')
    const indexEstado = buscarIndex('ESTADO')
    const colUser = excel.rows().getRows()

    for (let i = 0; i < colUser.length; i++) {

        for (let j = 0; j < colUser[i].length; j++) {

        }
    }

    console.log(colUser)
    console.log(colEstado)
    console.log(indexUser)
}

// Obtiene la fecha del archivo
function fechaReporte(excel) {
    const indexFecha = buscarIndex('FECHACREACION')
    const fechaCompleta = excel.columns().getCol(0).col[indexFecha].toString();
    const formateada = fechaCompleta.slice(0, fechaCompleta.indexOf(' '));
    document.querySelector("h5.fecha strong").innerText = formateada;
    // Obtiene la información de la columna correspondiente a 'FECHACREACION'.
    // Extrae la parte de la fecha hasta el primer espacio.
    // Actualiza el elemento HTML con la fecha formateada.
}


// muestra la informacion en este caso crear los botones con los nombres de cada usuario
class ExcelPrinter {
    static print(tableId, headLinksId, excel) {
        const table = document.getElementById(tableId)
        const headLinks = document.getElementById(headLinksId)

        const personal = arraySinRepetidos;

        personal.forEach((User, index) => {
            headLinks.querySelector("div").innerHTML += `<button class="btn-Links" id="${index}" onclick="userSelected(this.id)">${User}</button>`
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