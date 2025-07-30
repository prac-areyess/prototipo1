const puppeteer = require('puppeteer')

const Exceljs = require('exceljs')

const fs = require('fs').promises

//Modulo para llama archivos del sistema
const os = require('os')

const path = require('path')

let NombreResultadoCarpeta;

async function ConsultaPredial() {

    let browser
    try {

        const exectutableChromeNavegador = await getChromeExecutablePath()

        browser = await puppeteer.launch({
            headless: true, //Permite que el navegador sea visible
            slowMo: 100, //Simulacro la escritura
            defaultViewport: false,
            args: [
                '--start-maximized', //Permite utilizar el navegador en pantalla completa
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-infobars',
                '--disable-blink-features=AutomationControlled'
            ],
            executablePath: exectutableChromeNavegador
        })

        //Permite ejecutar el navegador nuevamente a partir de un cierre inesperado
        if (!browser || !browser.isConnected()) {
            browser = await puppeteer.launch();
            browser.on('disconnected', async () => {
                console.log('El navegador se ha cerrado. Reiniciando...');
                await ConsultaPredial();
            });
        }

        const page = await browser.newPage()

        // Enable CDP
        const client = await page.target().createCDPSession();

        const rutaDescargarDocumentos = path.join(process.cwd(), NombreResultadoCarpeta)

        const downloadPath = path.resolve(rutaDescargarDocumentos);

        await client.send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath
        });

        //const RutaActual = path.resolve(__dirname)

        const RutaUsuario = path.resolve(process.cwd(), 'Credenciales', 'usuario.txt')

        const RutaContraseña = path.resolve(process.cwd(), 'Credenciales', 'contrasena.txt')

        const contenidoUsuario = await loadUserAccountFile(RutaUsuario)

        const contenidioContraseña = await loadUserAccountFile(RutaContraseña)

        const RutaArchivoExcel = path.resolve(process.cwd(), 'Data', 'DATA_PERSONAS_JURIDICAS.xlsx')

        console.log('************ APP - SUNARP PJ ************')

        await page.goto('https://sprl.sunarp.gob.pe/sprl/ingreso', { waitUntil: "networkidle0", timeout: 300000 })

        let sidSunarp

        try {
            sidSunarp = await page.waitForSelector('svg[data-icon="close"]', { waitUntil: "networkidle0", timeout: 20000 })

            //  sidSunarp.click()
        } catch (error) {
            console.log('')
            await page.keyboard.press('Escape')
        }

        if (sidSunarp) {
            // await sidSunarp.click()
            // await page.$eval('button[aria-label="Close"]', elemento => elemento.click())
            await page.keyboard.press('Escape')
        }

        const btnIngresar = await page.waitForSelector('nz-form-control > div > div > div > button', { waitUntil: "networkidle0", timeout: 300000 })

        await btnIngresar.click()

        console.log(`Digitando Usuario...`)

        const usuario = await page.waitForSelector('input[name="username"]', { waitUntil: 'networkidle0', timeout: 300000 })

        await usuario.type(contenidoUsuario)

        console.log(`Digitando Contraseña...`)

        const contraseña = await page.waitForSelector('input[name="password"]', { waitUntil: 'networkidle0', timeout: 300000 })

        await contraseña.type(contenidioContraseña)

        console.log(`Iniciando Sesión...`)

        const btnInciarSesion = await page.waitForSelector('button[class="btn"]', { waitUntil: 'networkidle0', timeout: 300000 })

        await btnInciarSesion.click()

        // await page.waitForSelector('::-p-xpath(//app-main/nz-layout/nz-layout/nz-sider/div/app-sidenav-menu2/ul/li[1]/div[2]/ul/li/span[contains(text(),"partida")])', { waitUntil: 'networkidle0', timeout: 300000 })

        // await page.$eval('::-p-xpath(//app-main/nz-layout/nz-layout/nz-sider/div/app-sidenav-menu2/ul/li[1]/div[2]/ul/li/span[contains(text(),"partida")])', elem => elem.click())

        //const btnDesplegable = await page.waitForSelector('body > app-root > app-main > nz-layout > nz-header > div > div.header_A.ant-col > button',{waitUntil:'networkidle0', timeout:300000})

        //await btnDesplegable.click()

        // const consultaCertificada = await page.waitForSelector('div.ng-tns-c66-17.ant-menu.ant-menu-inline.ant-menu-sub.ng-trigger.ng-trigger-collapseMotion.ng-star-inserted > ul > li:nth-child(1)', {waitUntil:'networkidle0', timeout:300000})

        await page.waitForSelector('::-p-xpath(//app-main/nz-layout/nz-layout/nz-sider/div/app-sidenav-menu2/ul/li[1]/div[2]/ul/li/span[contains(text(),"partida")])', { waitUntil: 'networkidle0', timeout: 300000 })

        // const consultaCertificada = await page.waitForSelector('::-p-xpath(//app-main/nz-layout/nz-layout/nz-sider/div/app-sidenav-menu2/ul/li[1]/div[2]/ul/li[1])', {waitUntil:'networkidle0', timeout:300000})

        //const consultaCertificada = await page.waitForSelector('div.ng-tns-c66-15.ant-menu.ant-menu-inline.ant-menu-sub.ng-trigger.ng-trigger-collapseMotion.ng-star-inserted > ul > li:nth-child(1)',{waitUntil:'networkidle0', timeout:300000})

        //await page.$eval("div.ng-tns-c66-17.ant-menu.ant-menu-inline.ant-menu-sub.ng-trigger.ng-trigger-collapseMotion.ng-star-inserted > ul > li:nth-child(1)", elem => elem.click());

        await page.$eval('::-p-xpath(//app-main/nz-layout/nz-layout/nz-sider/div/app-sidenav-menu2/ul/li[1]/div[2]/ul/li/span[contains(text(),"partida")])', elemento => elemento.click())

        // await consultaCertificada.click()

        const NumeroTotalFilas = await FilasTotalesExcel(RutaArchivoExcel)

        //console.log(`Número total de Filas ${NumeroTotalFilas}`)

        for (let indice = 2; indice <= NumeroTotalFilas; indice++) {

            console.log(`----> Número de Consulta ${indice - 1} de ${NumeroTotalFilas - 1} <----`)

            let valorNumeroRUC = await CapturarValorExcel(RutaArchivoExcel, indice, 'A')
            let valorOficinaRegistral = await CapturarValorExcel(RutaArchivoExcel, indice, 'B')
            let valorNumeroPartida = await CapturarValorExcel(RutaArchivoExcel, indice, 'C')
            let estadoConsulta = await CapturarValorExcel(RutaArchivoExcel, indice, 'D')

            valorNumeroRUC = String(valorNumeroRUC).toLowerCase()
            valorOficinaRegistral = String(valorOficinaRegistral).toLowerCase()
            valorNumeroPartida = String(valorNumeroPartida).toLowerCase()
            estadoConsulta = String(estadoConsulta).toLowerCase()

            if (estadoConsulta.includes("encontrado") || estadoConsulta.includes("no encontrado")) {
                console.log('-> Siguiente consulta')
                console.clear()
            } else {

                const OficinaRegistral = await page.waitForSelector('::-p-xpath(//nz-form-item/nz-form-control/div/div/nz-select/nz-select-top-control/nz-select-search/input)', { waitUntil: 'networkidle0', timeout: 300000 })

                //await OficinaRegistral.type('LIMA')
                await OficinaRegistral.type(valorOficinaRegistral)

                await page.keyboard.press('Enter')

                const servicioPredial = await page.waitForSelector("nz-select-item[title='FIR.DIGITAL-CERT. LITERAL - PREDIOS']")

                await servicioPredial.click()

                const servicioPersonasJuridicas = await page.waitForSelector("nz-option-item[title='FIR.DIGITAL-CERT. LITERAL - PJ']")

                await servicioPersonasJuridicas.click()

                //const btnPartida = await page.waitForSelector('input[type="radio"]',{waitUntil:'networkidle0'})

                //await btnPartida.click()

                await page.$eval("input[type='radio']", elem => elem.click());

                const campoNumero = await page.waitForSelector('input[name="numero"]')

                await campoNumero.type(valorNumeroPartida)

                await page.keyboard.press('Enter')

                await page.$eval('button[type="submit"]', elem => elem.click());

                //let Bandera = false 
                let boleeanPartidaNoEncontrada
                let boleeanSolictiudDescarga

                try {
                    const [partidaNoEncontrada, solicitudDescarga] = await Promise.all([
                        (async () => {
                            try {
                                await page.waitForSelector('::-p-xpath(//nz-modal-confirm-container/div/div/div/div/div[1]/span/span)', { visible: true, timeout: 15000 });
                                page.reload();
                                // const solicitarCertificadoRefresh = await page.waitForSelector('::-p-xpath(//app-main/nz-layout/nz-layout/nz-sider/div/app-sidenav-menu2/ul/li[1]/div[2]/ul/li[1])', { waitUntil: 'networkidle0', timeout: 300000 });
                                // await solicitarCertificadoRefresh.click();

                                await page.waitForSelector('::-p-xpath(//app-main/nz-layout/nz-layout/nz-sider/div/app-sidenav-menu2/ul/li[1]/div[2]/ul/li/span[contains(text(),"partida")])', { waitUntil: 'networkidle0', timeout: 300000 })

                                await page.$eval('::-p-xpath(//app-main/nz-layout/nz-layout/nz-sider/div/app-sidenav-menu2/ul/li[1]/div[2]/ul/li/span[contains(text(),"partida")])', elemento => elemento.click())

                                console.log('--> Número Partida No Encontrada');
                                await EscribirArchivoExcel(RutaArchivoExcel, indice, 'D', 'no encontrado');
                                return true;
                            } catch (error) {
                                return false;
                            }
                        })(),
                        (async () => {
                            try {
                                const btnSolicitud = await page.waitForSelector('nz-content > div:nth-child(6) > button.ant-btn.ant-btn-primary', { timeout: 15000 });
                                await btnSolicitud.click();
                                return true;
                            } catch (error) {
                                return false;
                            }
                        })()
                    ]);

                    boleeanPartidaNoEncontrada = partidaNoEncontrada;
                    boleeanSolictiudDescarga = solicitudDescarga;

                    //console.log('-> Partida No Encontrada:', boleeanPartidaNoEncontrada);
                    //console.log('-> Solicitud Descarga:', boleeanSolictiudDescarga);

                } catch (error) {
                    console.error('Error ejecutando las consultas en paralelo:', error);
                }


                if (boleeanPartidaNoEncontrada == false || boleeanSolictiudDescarga == true) {

                    const btnVerAsientos = await page.waitForSelector("button[title='Ver Asientos'")

                    await btnVerAsientos.click()

                    const btnTodasLasPaginas = await page.waitForSelector('nz-radio-group > label:nth-child(1) > span.ant-radio > input', { waitUntil: 'networkidle0', timeout: 300000 })

                    await btnTodasLasPaginas.click()

                    const btnCalcularMonto = await page.waitForSelector('app-consulta-partidas > nz-content > nz-spin > div > app-ver-asientos > div.montos > div:nth-child(3) > button', { waitUntil: 'networkidle0', timeout: 300000 })

                    await btnCalcularMonto.click()

                    if (boleeanSolictiudDescarga == true) {

                        const btnSaldoDisponible = await page.waitForSelector(' app-ver-asientos > app-radio-buttom-custom > div > nz-form-item > nz-form-control > div > div > div > nz-radio-group > label:nth-child(2) > span.ant-radio > input')

                        await btnSaldoDisponible.click()

                    }

                    const btnContinuar = await page.waitForSelector('app-ver-asientos > app-button-triple2 > div > div:nth-child(2) > app-button > div > div > button', { waitUntil: 'networkidle0', timeout: 300000 })

                    await btnContinuar.click()

                    const btnDescargarBoleta = await page.waitForSelector("button[class='ant-btn ant-btn-primary']", { waitUntil: 'networkidle0', timeout: 300000 })

                    await btnDescargarBoleta.click()

                    console.log(`--> Descargando La Partida Registral ${valorNumeroPartida}`)

                    // Monitor download progress
                    client.on('Page.downloadProgress', (event) => {
                        if (event.state === 'completed') {
                            console.log('Download completed');
                        }
                    });

                    //await DescargarBoletaInformativa()

                    await RenombrarDocumento(rutaDescargarDocumentos, `${indice - 1}_${valorNumeroRUC}_${valorNumeroPartida}`, `PARTIDA ${valorNumeroPartida}.pdf`)

                    await EscribirArchivoExcel(RutaArchivoExcel, indice, 'D', 'encontrado')

                    const btnRegresar = await page.waitForSelector("button[class='not-print ant-btn ant-btn-primary']", { waitUntil: 'networkidle0', timeout: 18000 })

                    await btnRegresar.click()

                }

                await CopiaryMoverArchivoExcel(RutaArchivoExcel, rutaDescargarDocumentos, `${NombreResultadoCarpeta}.xlsx`)

            }

            if (indice == NumeroTotalFilas) {
                console.log('---> El Proceso a Finalizado Con Éxito <---')
                await browser.close()
            }

            //Cuidado
        }

    } catch (error) {
        console.error('Error al iniciar el navegador o realizar consultas:', error);
        //console.clear();
        // Cierra el navegador si hay un error
        if (browser) {
            await browser.close();
        }
        // Reintenta después de 15 segundos
        setTimeout(ConsultaPredial, 15000);
    }

}

async function FilasTotalesExcel(ArchivoExcel) {

    const workbook = new Exceljs.Workbook();

    await workbook.xlsx.readFile(ArchivoExcel)

    let obtenerNumeroFilas = 0;

    const worksheet = workbook.getWorksheet(1);

    worksheet.eachRow((row) => {

        if (row.hasValues) {
            obtenerNumeroFilas++;
        }
    })

    return obtenerNumeroFilas
}

async function CapturarValorExcel(ArchivoExcel, indiceFila, LetraVertical) {

    const workbook = new Exceljs.Workbook();

    await workbook.xlsx.readFile(ArchivoExcel)

    const worksheet = workbook.getWorksheet(1);

    return worksheet.getCell(`${LetraVertical}${indiceFila}`).value

}

async function EscribirArchivoExcel(ArchivoExcel, indice, LetraVertical, texto) {

    // Crear un nuevo libro de trabajo
    const workbook = new Exceljs.Workbook();

    // Leer el archivo existente
    await workbook.xlsx.readFile(ArchivoExcel);

    // Seleccionar la primera hoja de trabajo
    const worksheet = workbook.getWorksheet(1);

    // Escribir datos en la celda especificada
    worksheet.getCell(`${LetraVertical}${indice}`).value = texto;

    // Guardar el archivo
    await workbook.xlsx.writeFile(ArchivoExcel);
}


async function RenombrarDocumento(rutaActual, nuevaRuta, nuevoNombre) {
    try {
        let estado = false
        let tiempoAcumulado = 0 //segundos de espera
        let tiempoEspera = 5000

        while (estado == false) {

            await new Promise(resolve => setTimeout(resolve, tiempoEspera));
            tiempoAcumulado += tiempoEspera

            // Leer el contenido de la carpeta
            const archivos = await fs.readdir(rutaActual);

            // Encontrar el archivo PDF
            const archivoPDF = archivos.find(archivo => path.extname(archivo).toLowerCase() === '.pdf');

            if (archivoPDF) {
                const rutaActualCompleta = path.join(rutaActual, archivoPDF);
                const nuevaRutaCompleta = path.join(rutaActual, nuevaRuta, nuevoNombre);

                // Crear el directorio de destino si no existe
                await fs.mkdir(path.dirname(nuevaRutaCompleta), { recursive: true });

                // Mover y renombrar el archivo
                await fs.rename(rutaActualCompleta, nuevaRutaCompleta);
                console.log('---> Partida Registral Descargada Con Éxito')
                //console.log(`---> Documento ${archivoPDF} renombrado y movido a ${nuevaRutaCompleta}`);
                estado = true
            } else {
                console.log('---> Descargando ...')
                //console.log('---> No se encontró ningún archivo PDF en la carpeta especificada.');
            } if (tiempoAcumulado >= 180000) { // 3 minutos
                throw new Error('Tiempo de espera acumulado excedido. Cerrando el navegador.');
            }
        }
    } catch (error) {
        console.error('---> Error al renombrar y mover el documento:', error);
        //console.clear();
        // Cierra el navegador si hay un error
        if (browser) {
            await browser.close();
        }
        // Reintenta después de 15 segundos
        setTimeout(ConsultaPredial, 15000);
    }
}

async function FormatoFecha() {
    const fechaActual = new Date()

    const hostName = os.hostname().toLocaleUpperCase()
    const userInfo = os.userInfo().username.toLocaleUpperCase()

    let diaActual = fechaActual.getDate()
    let mesActual = fechaActual.getMonth() + 1
    let anioActual = fechaActual.getFullYear()
    let horaActual = fechaActual.getHours()
    let minutosActuales = fechaActual.getMinutes()
    let segundosActuales = fechaActual.getSeconds()

    //console.log(mesActual)

    if (diaActual.toString().length === 1) {
        diaActual = `0${fechaActual.getDate()}D`
    } else {
        diaActual = fechaActual.getDate() + 'D'
    }

    if (mesActual.toString().length === 1) {
        mesActual = `0${fechaActual.getMonth() + 1}M`
    } else {
        mesActual = `${fechaActual.getMonth() + 1}M`
    }

    if (horaActual.toString().length === 1) {
        horaActual = `0${fechaActual.getHours()}H`
    } else {
        horaActual = fechaActual.getHours() + 'H'
    }

    if (minutosActuales.toString().length === 1) {
        minutosActuales = `0${fechaActual.getMinutes()}M`
    } else {
        minutosActuales = `${fechaActual.getMinutes()}M`
    }

    if (segundosActuales.toString().length === 1) {
        segundosActuales = `0${fechaActual.getSeconds()}S`
    } else {
        segundosActuales = fechaActual.getSeconds() + 'S'
    }

    return `RESULTADO_SUNARP_PJ_${userInfo}_${diaActual}_${mesActual}_${anioActual}_${horaActual}_${minutosActuales}_${segundosActuales}`
}

async function CrearDirectorioDescargas() {
    NombreResultadoCarpeta = await FormatoFecha()
    const rutaDescargarDocumentos = path.join(process.cwd(), NombreResultadoCarpeta)
    try {
        await fs.mkdir(rutaDescargarDocumentos, { recursive: true });
        //console.log(`Directorio creado exitosamente ${rutaDescargarDocumentos}`);
    } catch (err) {
        console.log('Error al crear el directorio:', err);
    }
}


async function getChromeExecutablePath() {
    const pathA = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
    const pathB = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe';

    try {
        const pathAExists = await fs.access(pathA).then(() => true).catch(() => false);
        return pathAExists ? pathA : pathB;
    } catch (error) {
        console.error('Error al encontrar la ruta:', error);
        return null;
    }
}


async function loadUserAccountFile(archivoUsuario) {
    try {
        const textoUserAccount = await fs.readFile(archivoUsuario, 'utf-8');
        return textoUserAccount;
    } catch (error) {
        console.error('Error al leer el archivo:', error);
        return null;
    }
}


async function CopiaryMoverArchivoExcel(RutaBase, RutaDestino, NuevoNombre) {
    try {
        // Asegúrate de que el directorio de destino exista
        await fs.mkdir(RutaDestino, { recursive: true });

        // Define la ruta completa del archivo de destino
        //const fileName = path.basename(RutaBase);
        const destinationPath = path.join(RutaDestino, NuevoNombre);

        // Copia el archivo
        await fs.copyFile(RutaBase, destinationPath);

        //console.log(`Archivo copiado a: ${destinationPath}`);
    } catch (error) {
        console.error('Error al copiar y mover el archivo:', error);
    }
}



async function main() {
    await CrearDirectorioDescargas()
    await ConsultaPredial()
}

main()