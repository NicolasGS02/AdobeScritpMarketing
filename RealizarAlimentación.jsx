Main();

function Main() {

    //VARIABLES CONSTANTES


    //Variable donde se encontrarán las imagenes.
    const direccionImagenes = "D:\\Documentos\\Revista Papá\\IMAGENES CARDOSO ABRIL 2025";
    //Extensión de las imagenes permitidas.
    var formatos = ["jpg", "jpeg", "png", "gif", "pdf", "eps", "tif"];
    const marcas = ["VIVÓ","1890", "ALBO", "ALHAMBRA", "ALMIRANTE", "ALSUR", "ALTEZA", "AMSTEL", "APIS", "ARLUY",
        "ARTIACH", "ASTURIANA", "AVECREM", "AVIVA", "BAILEYS", "BAJAMAR", "BALEA", "BALLANTINES",
        "BANDEROS", "BARCELO", "BEEFEATER", "BIMBO", "BOFFARD", "BORGES", "BREKKIES", "BRILLANTE",
        "BRUNO", "BUCKLER", "BUDWEISER", "BUITONI", "CACAOLAT", "CACIQUE", "CALVE", "CALVO",
        "CAMPO", "CAMPOCURADO", "CAMPOFRIO", "CAMPOFRÍO", "CARBONELL", "CERRATO", "COCA", "COLA-CAO",
        "COVAP", "CRISMONA", "CRUZCAMPO", "CUTTY", "CUÉTARA", "DANONE", "DELAVIUDA", "DELEITUM",
        "DESPERADOS", "DON SIMON", "DONUTS", "EMPARRADO", "ESTRELLA", "FANTA", "FINDUS", "FIZZY",
        "FLOR", "FRAGATA", "FRESKITO", "FRUCO", "FRUDESA", "FUZE", "GALLO", "GAMITO", "GARCÍA",
        "GAROFALO", "GOODFELLA", "HEINEKEN", "HELLMANNS", "HUESITOS", "INÉS", "IZNAOLIVA",
        "JARLSBERG", "JOSÉ", "JUVER", "KAS", "KIKKOMAN", "KNORR", "LAFUENTE", "LEYENDA", "LIGERESA",
        "LITORAL", "LOCURA", "LORENZANA", "LOTAMAR", "LOURIÑO", "LUENGO", "LUXMAR", "LVF", "MAGGI",
        "MAHOU", "MALTESERS", "MAMMEN", "MARTINETE", "MASIASOL", "MATA", "MILKA", "MUCHO", "MUSA",
        "NAVIDUL", "NESCAFE", "NESCAFÉ", "NESQUIK", "NESTLE", "NESTLÉ", "NOVA", "OKEY", "ORLANDO",
        "OSCAR MAYER", "PASCUAL", "PAULANER", "PAY", "PEPSI", "PEÑA", "POK", "POPITAS", "PRIMA",
        "PROLONGO", "PUERTO", "PULEVA", "QUESCAN", "RAM", "REVILLA", "RIBEIRA", "RIBERA", "RIKISSSIMO",
        "RIOJA", "RIOVERDE", "ROSIL", "ROYAL", "RUEDA", "SAGRES", "SEDA", "SERRANO", "SEVEN",
        "SHANDY", "SOLIS", "TAKIS", "TEJERO", "TELLO", "TIERRA", "TULIPÁN", "UBAGO", "ULTIMA",
        "USISA", "VALDEPEÑA", "VALDEPEÑAS", "VALOR", "VIVO", "VIVÓ", "VODKA", "WHITE", "XOCOTOK",
        "YBARRA", "YOLANDA", "GALLINA BLANCA", "YZAGUIRRE", "LA COCINERA", "RIKISSSIMO"
    ];

    //ESTILOS
    var nombreEstiloTitulo = "textoTitulo";
    var nombreEstiloPrecio = "PVP_NEW";
    var nombreEstiloTextoLey = "LEY";


    //COORDENADAS DE LOS ELEMENTOS. POSICION DESTACADO > 9|| POSICION > 9 productos || POSICION DESTACADO || POSICION NORMAL
    var posicionImagen = [[14, 7, 92, 76], [7, 143, 58, 185], [15, 7, 83, 110], [103, 7, 150, 49]];
    var posicionesTexto = [[23, 76, 50, 136], [58, 143, 73, 185], [18, 106, 59, 185], [154, 7, 170, 49]];
    var posicionPrecios = [[55, 76, 70, 136], [73, 143, 88, 185], [59, 106, 79, 185], [170, 7, 185, 49]];
    var posicionTextoLey = [[70, 76, 75, 136], [88, 143, 96, 185], [79, 106, 96, 185], [185, 7, 191, 49]];


    var file = File.openDialog("Selecciona un archivo XLSX", "*.xlsx");
    //COnseguimos el Path del archivo para poder abrirlo.
    var excelFilePath = file.fsName;
    var splitChar = ";";

    //Esta variable nos dará el número de páginas que tiene el archivo de excel.
    var numPaginas = 2;


    //Desplazamiento de los productos en la página.
    var desplazamientoDerecha = 45.25;
    var desplazamientoAbajo = 92;

    //------------------------------------------------------

    //Este bucle hará que vayamos cargando poco a poco las paginas del documento Exel.
    for (var i = 1; i <= numPaginas; i++) {

        var datosRecogidosExel = GetDataFromExcelPC(excelFilePath, splitChar, i);

        var cantidadDeProductos = 0;
        var cont = 2; //Empezamos a contar desde la fila 2, ya que la fila 1 son los titulos de las columnas.
        var salir = true; //Variable que nos ayudará a salir del bucle.

        do {
            if (datosRecogidosExel[cont][1] != "") {
                cantidadDeProductos++;
                cont++;
            } else {
                salir = false;
            }
        } while (salir);




        //------------------------------------------------------
        //DOCUMENTO BASE DE INDESIGN.

        //Creamos un nuevo documento de indesign.
        var doc = app.activeDocument;
        //Crearemos una nueva pagina en el documento de indesign.
        var nuevaPagina = doc.pages.add(LocationOptions.AFTER, doc.pages[doc.pages.length - 1]);

        //------------------------------------------------------
        //ANALISIS COLUMNAS

        //Primero miraremos en que posición se encuentran las columnas que nos interesan.
        //En nuestro caso, en la variable datosRecogidosExel[0] tenemos los titulos de las columnas. Y buscaremos las que digan:ç
        //"CÓDIGO" para la columna de imagenes.
        //"DESCRIPCIÓN" para la columna de nombre.
        //"P.V.O." para la columna de precio.
        //"TEXTO LEY" para la columna de texto ley.

        //Creamos variables donde guardaremos los indices de las columnas que nos interesan. 
        var columnaImagen = 0;
        var columnaNombre = 0;
        var columnaPrecio = 0;
        var columnaTextoLey = 0;

        for (var j = 0; j < datosRecogidosExel[0].length; j++) {

            if (datosRecogidosExel[0][j].toUpperCase() == "CODIGO") {
                columnaImagen = j;
            } else if (datosRecogidosExel[0][j].toUpperCase() == "DESCRIPCION") {
                columnaNombre = j;
            } else if (datosRecogidosExel[0][j].toUpperCase() == "P.V.O.") {
                columnaPrecio = j;
            } else if (datosRecogidosExel[0][j].toUpperCase() == "TEXTO LEY") {
                columnaTextoLey = j;
            }
        }



        //Para quitarnos las filas que no tienen datos, creamos un nuevo array con los datos que nos interesan.
        datosRecogidosExel = datosRecogidosExel.splice(2, cantidadDeProductos); //Quitamos las dos primeras filas, ya que son los titulos de las columnas y la fila 2 que no tiene datos.
        var contadorPosicion = 0;
        //BUCLE PRINCIPAL PARA ITERAR POR LOS PRODUCTOS.
        for (var j = 0; j < cantidadDeProductos; j++) {

            var imagen = datosRecogidosExel[j][columnaImagen];
            var nombre = datosRecogidosExel[j][columnaNombre];
            var precio = datosRecogidosExel[j][columnaPrecio];
            var textoLey = datosRecogidosExel[j][columnaTextoLey];


            //------------------------------------------------------
            //TRATAMIENTO DE LAS IMAGENES.

            //Primero miramos cuantas codigos de imagenes tenemos.
            var numeros = imagen.split("-");
            //Creamos un array para guardar los archivos de imagenes.
            var imagenArchivo = new Array(numeros.length);


            //Recorremos el array de imagenes y vamos buscando los archivos de imagenes en la carpetas.
            for (var k = 0; k < numeros.length; k++) {
                for (var y = 0; y < formatos.length; y++) {
                    //Aqui buscamos si existe el archivo con las extensiones que hemos definido antes. K=Indices de productos y Y=Indices de formatos.
                    var imagenBuscada = new File(direccionImagenes + "/" + numeros[k] + "." + formatos[y]);
                    if (imagenBuscada.exists) {
                        imagenArchivo[k] = imagenBuscada;
                        break; //Salimos del bucle si hemos encontrado la imagen.
                    }
                }
            }
            //------------------------------------------------------
            //PONER LOS DATOS EN EL DOCUMENTO DE INDESIGN.
            var textoProducto = nuevaPagina.textFrames.add();
            var textoPrecio = nuevaPagina.textFrames.add();
            var textoPrecioLey = nuevaPagina.textFrames.add();

            //Ahora debemos diferenciar si son 9 productos o más de eso, para ver la distrobución de los productos en la página.

            //Comprobamos que los destacados solamente se colocan en la primera iteración del bucle de productos.
            if (cantidadDeProductos > 9 && j < 2) {

                //primero deberemos colocar el producto destacado y el producto extra.
                //Colocamos el producto destacado.

                //Creamos los textFrames para el nombre, precio y texto ley.
                //Cuando p=0, colocamos el producto destacado y cuando p=1, colocamos el producto extra.



                //Posicionamos imagen:

                //Como puede haber más de 1 imagen,tenemos que ver cuantas hay.
                for (var im = 0; im < imagenArchivo.length; im++) {

                    if (imagenArchivo[im] != undefined) {
                        var imagen = nuevaPagina.rectangles.add();
                        imagen.strokeWeight = 0;
                        imagen.geometricBounds = posicionImagen[j];
                        imagen.place(imagenArchivo[im]);
                        imagen.fit(FitOptions.PROPORTIONALLY);

                    }

                }

                //Colocamos el texto del nombre del producto.
                textoProducto.geometricBounds = posicionesTexto[j];
                nombre = insertarSaltoTrasMarca(nombre, marcas);
                textoProducto.contents = nombre;

                //Colocamos el texto del precio.
                textoPrecio.geometricBounds = posicionPrecios[j];
                textoPrecio.contents = precio;

                //Colocamos el texto del texto ley.
                textoPrecioLey.geometricBounds = posicionTextoLey[j];
                textoPrecioLey.contents = textoLey;







            } else {
                //Ahora vamos a colocar los productos normales.

                //Dependiendo la posición del producto, nos tendremos que desplazar a la derecha, y en algún momenos hacia abajo.

                if (imagenArchivo.length > 0) {
                    //Primero comprueba si hace falta desplazar hacia abajo, si es mayor a 4, bajamos una fila.
                    var desplazamientoVertical = (contadorPosicion < 4) ? 0 : desplazamientoAbajo;
                    //Ahora comporbamos para ir abanzando hacia la derecha
                    var desplazamientoHorizontal = desplazamientoDerecha * ((contadorPosicion < 4) ? contadorPosicion : (contadorPosicion - 4));

                    for (var im = 0; im < imagenArchivo.length; im++) {
                        if (imagenArchivo[im] !== undefined) {
                            var imagen = nuevaPagina.rectangles.add();
                            imagen.strokeWeight = 0;
                            imagen.geometricBounds = [
                                posicionImagen[3][0] + desplazamientoVertical,
                                posicionImagen[3][1] + desplazamientoHorizontal,
                                posicionImagen[3][2] + desplazamientoVertical,
                                posicionImagen[3][3] + desplazamientoHorizontal
                            ];
                            imagen.place(imagenArchivo[im]);
                            imagen.fit(FitOptions.PROPORTIONALLY);
                        }
                    }


                    //Colocamos el texto del nombre del producto.
                    textoProducto.geometricBounds = [
                        posicionesTexto[3][0] + desplazamientoVertical,
                        posicionesTexto[3][1] + desplazamientoHorizontal,
                        posicionesTexto[3][2] + desplazamientoVertical,
                        posicionesTexto[3][3] + desplazamientoHorizontal
                    ];
                    //Si el nombre del producto tiene alguna de las marcas, le añadimos un salto de línea.
                    nombre = insertarSaltoTrasMarca(nombre, marcas);
                    textoProducto.contents = nombre;

                    //Colocamos el texto del precio.
                    textoPrecio.geometricBounds = [
                        posicionPrecios[3][0] + desplazamientoVertical,
                        posicionPrecios[3][1] + desplazamientoHorizontal,
                        posicionPrecios[3][2] + desplazamientoVertical,
                        posicionPrecios[3][3] + desplazamientoHorizontal
                    ];
                    //Añadimos el símbolo de euro al precio.
                    textoPrecio.contents = precio + "€";

                    //Colocamos el texto del texto ley.
                    textoPrecioLey.geometricBounds = [
                        posicionTextoLey[3][0] + desplazamientoVertical,
                        posicionTextoLey[3][1] + desplazamientoHorizontal,
                        posicionTextoLey[3][2] + desplazamientoVertical,
                        posicionTextoLey[3][3] + desplazamientoHorizontal
                    ];
                    textoPrecioLey.contents = textoLey;








                    contadorPosicion++;
                }
            }


                    //Ahora vamos a aplcar el formato al texto.
                    //Primero el texto del nombre del producto, que es algo más complejo.


                    // Estilo inicial
                    var firstStyle = app.activeDocument.paragraphStyles.itemByName(nombreEstiloTitulo);

                    var currentStyle = firstStyle;

                    // Aplicamos primero el estilo inicial
                    var paragraphs = textoProducto.paragraphs;
                    for (var li = 0; li < paragraphs.length; li++) {
                        paragraphs[li].applyParagraphStyle(currentStyle, true);

                        // Si el estilo tiene definido un "estilo siguiente", lo cogemos
                        if (currentStyle.nextStyle != null && currentStyle.nextStyle.isValid) {
                            currentStyle = currentStyle.nextStyle;
                        }
                    }
                

                    //Después los del precio y el texto ley.
                    var estiloPrecio = doc.paragraphStyles.item(nombreEstiloPrecio);
                    var estiloTextoLey = doc.paragraphStyles.item(nombreEstiloTextoLey);


                    textoPrecio.paragraphs[0].applyParagraphStyle(estiloPrecio);
                    textoPrecioLey.paragraphs[0].applyParagraphStyle(estiloTextoLey);




        }//Fin del bucle de productos.

    }//Fin del bucle de paginas.


}

function insertarSaltoTrasMarca(nombre, marcas) {
    for (var k = 0; k < marcas.length; k++) {
        var marca = marcas[k];
        // Escapar caracteres especiales en la marca
        var safeMarca = marca.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
        var regex = new RegExp("\\b(" + safeMarca + ")\\b\\s*", "i");

        if (regex.test(nombre)) {
            // Sustituye por la marca + salto de párrafo sin espacio
            nombre = nombre.replace(regex, "$1\r");
            break;
        }
    }
    return nombre;
}




function GetDataFromExcelPC(excelFilePath, splitChar, sheetNumber) {
    if (typeof splitChar === "undefined") var splitChar = ";";
    if (typeof sheetNumber === "undefined") var sheetNumber = "1";
    var appVersionNum = Number(String(app.version).split(".")[0]);

    var vbs = 'Public s\r';
    vbs += 'Function ReadFromExcel()\r';
    vbs += 'Set objExcel = CreateObject("Excel.Application")\r';
    vbs += 'Set objBook = objExcel.Workbooks.Open("' + excelFilePath + '")\r';
    vbs += 'Set objSheet =  objExcel.ActiveWorkbook.WorkSheets(' + sheetNumber + ')\r';
    vbs += 'objExcel.Visible = False\r';
    vbs += 'matrix = objSheet.UsedRange\r';
    vbs += 'maxDim0 = UBound(matrix, 1)\r';
    vbs += 'maxDim1 = UBound(matrix, 2)\r';
    vbs += 'For i = 1 To maxDim0\r';
    vbs += 'If i > 20 Then Exit For\r'; // Limit to first 20 rows
    vbs += 'For j = 1 To maxDim1\r';
    vbs += 'If j = maxDim1 Then\r';
    vbs += 's = s & matrix(i, j)\r';
    vbs += 'Else\r';
    vbs += 's = s & matrix(i, j) & "' + splitChar + '"\r';
    vbs += 'End If\r';
    vbs += 'Next\r';
    vbs += 's = s & vbCr\r';
    vbs += 'Next\r';
    vbs += 'objBook.Close\r';
    vbs += 'Set objSheet = Nothing\r';
    vbs += 'Set objBook = Nothing\r';
    vbs += 'Set objExcel = Nothing\r';
    vbs += 'End Function\r';
    vbs += 'Function SetArgValue()\r';
    vbs += 'Set objInDesign = CreateObject("InDesign.Application")\r';
    vbs += 'objInDesign.ScriptArgs.SetValue "excelData", s\r';
    vbs += 'End Function\r';
    vbs += 'ReadFromExcel()\r';
    vbs += 'SetArgValue()\r';

    if (appVersionNum > 5) { // CS4 and above
        app.doScript(vbs, ScriptLanguage.VISUAL_BASIC, undefined, UndoModes.FAST_ENTIRE_SCRIPT);
    }
    else { // CS3 and below
        app.doScript(vbs, ScriptLanguage.VISUAL_BASIC);
    }

    var str = app.scriptArgs.getValue("excelData");
    app.scriptArgs.clear();

    var tempArrLine, line,
        data = [],
        tempArrData = str.split("\r");

    for (var i = 0; i < tempArrData.length; i++) {
        line = tempArrData[i];
        if (line == "") continue;
        tempArrLine = line.split(splitChar);
        data.push(tempArrLine);
    }

    return data;
}