const express = require('express');
const bodyParser = require("body-parser");

//Ver cual es el que funciona y deshacte del que no.
const pdfDocument = require("pdfkit");
const XlsxPopulate = require("xlsx-populate")

// //Convertir Excel a pdf
// const puppeteer = require("puppeteer")
// const xslx = require("xlsx")
// const fs = require("fs")

// const pdfService = require("./micro-servico/buildPDF")

const app = express()
const port = 3000

// const handlebars = require("express-handlebars");

app.set("view engine", "hbs");



// app.engine("hbs", handlebars.engine({
//     layoutsDir: `${__dirname}/views/layouts`,
//     extname: "hbs",

//     //custom helper
//     helpers: {

//         objToList : function(context) {
//             function toList(obj, indent) {
//               var res=""
//               for (var k in obj) { 
//                   if (obj[k] instanceof Object) {
//                       res=res+k+"\n"+toList(obj[k], ("" + indent)) ;
//                   }
//                   else{
//                       res=res+indent+k+" : "+obj[k]+"\n";
//                   }
//               }
//               return res;
//             }    
//             return toList(context,"");
//         }
//     }
// }));

app.use(express.static("public"));


app.use(express.json())

//Middleware
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: false}));


app.post("/", (req, res) => {
    const body = req.body;


    res.render("main", {layout:"index", body})
})

app.post("/generate-pdf", (req, res) => {
    try {
        const body = req.body;
        res.setHeader("Content-Type", "application/pdf");
        // pdfkit.create(html, options).pipe(res);
        res.render("main", {layout:"index", body}, (err, html) => {
            if (err) {
                console.error(err);
                res.status(500).send(err);
            } else {
                function buildPDF() {
                  const doc = new pdfDocument({
                    size: "A4",
                    layout: "portrait"
                  });
                  doc.pipe(res);
                  doc.fontSize(25).text(html);
                  doc.end();
                }
                buildPDF();
            }
        });
    } catch (err) {
        console.error(err);
        res.status(500).send(err);
    }
});




app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})


app.post("/generate-excel", (req, res) => {
    try {
        const data = req.body;
        XlsxPopulate.fromFileAsync("./template/template.xlsx")
            .then(workbook => {
            let r = 11
            let startRow = r + 1
            //CURP
            workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("curp").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
            workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(data.curp).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});

            // Trabajo Activo
            r = 13
            workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Trabajo Activo").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
            if (data.imss?.data?.historialLaboral[0]?.fechaBaja === "Vigente" ||
                data.issste?.data?.personalInfo?.SituacionAfiliatoria === "ACTIVO") {
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value("True").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value("False").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"FF0000"});
            }
            //Resultados
            r = 15
            workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Resultados").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
            workbook.sheet("Reporte de ingresos").cell(`C${r}`).value("Información Encontrada").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});

            //  IMSS
            r = 16
            workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("IMSS").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
            if (data.imss.status === "ERROR") {
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value("false").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"FF0000"});;
            } else {
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value("true").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
            }
            workbook.sheet("Reporte de ingresos").cell(`D${r}`).value(data.imss.message).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});

            // ISSSTE
            r = 17
            workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("ISSSTE").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
            if (data.issste.data.message === "CURP inválido o no encontrado en el ISSSTE") {
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value("false").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"FF0000"});;
                workbook.sheet("Reporte de ingresos").cell(`D${r}`).value(data.issste.data.message).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
            } else {
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value("true").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`D${r}`).value("Información Encontrada").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
            }

            // IMSS Detail
            r = 19
            workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Información del IMSS").style({"bold" : true, "fontSize": 20, "fontColor": "2F5496"}) //INFORMACION DEL IMSS
            r = 20
            if (data.imss.status !== "ERROR" && data.imss) {
                
                //NSS
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("NSS").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(data.imss.nss).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})
            
                //Nombre
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Nombre").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(data.imss.data.nombre).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Semanas Cotizadas").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(data.imss.data.semanasCotizadas.semanasCotizadas).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Semanas Reintegradas").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(data.imss.data.semanasCotizadas.semanasReintegradas).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Semanas Reitengradas").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(data.imss.data.semanasCotizadas.semanasDescontadas).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})

                r++
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Historial Laboral").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                data.imss.data.historialLaboral.forEach(el => {
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Fecha Alta").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.fechaAlta).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Fecha Baja").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.fechaBaja).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Antiguedad").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.antiguedad).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Salario Base Cotización").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.salarioBaseCotizacion).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Salario Mensual").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.salarioMensual).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Nombre Patrón").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.nombrePatron).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Registro Patronal").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.registroPatronal).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                r++
                
                //Loop para darle color al background
                for (let i = 20; i <= r; i++) {
                    if (i % 2 === 0) {
                      workbook.sheet("Reporte de ingresos").range(`B${i}:C${i}`).style({
                        "fill": {
                          type: "pattern",
                          pattern: "solid",
                          foreground: {
                            rgb: "ffffff"
                          },
                          background: {
                            theme: 3,
                            tint: 0.4
                          }
                        }
                      });
                    } else {
                      workbook.sheet("Reporte de ingresos").range(`B${i}:C${i}`).style({
                        "fill": {
                          type: "pattern",
                          pattern: "solid",
                          foreground: {
                            rgb: "F0F0F0"
                          },
                          background: {
                            theme: 3,
                            tint: 0.4
                          }
                        }
                      });
                    }
                  }
                })

            } else {
                // No se encontró historial del IMSS
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value(data.imss.message);
                r++
            }

            // ISSSTE
            r++
            r++
            r++
            workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Información del ISSSTE").style({"bold" : true, "fontSize": 20, "fontColor": "2F5496"})//INFORMACION DEL ISSTE
            startRow = r + 1
            r++
            r++
            if(data.issste?.data?.message !== "CURP inválido o no encontrado en el ISSSTE"){
                const x = data.issste?.data
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Nombre").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.Nombre).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Primer Apellido").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.primerApellido).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Segundo Apellido").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.segundoApellido).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Segundo Apellido").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.segundoApellido).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("RFC").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.RFC).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Sexo").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.Sexo);
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Fecha de Nacimiento").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.fechaDeNacimiento).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Estado Civil").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.fechaDeNacimiento).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Número de seguridad social").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.numeroSeguridadSocial).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Fecha de alta en la plaza actual").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.fechaAltaPlazaActual).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Fecha de ingreso al Gobierno Federal").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.fechaIngresoAlGobiernoFederal).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Fecha de baja").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.fechaDeBaja).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Derecho a servicio médico").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.DerechoServicioMedico).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Tipo de derechohabiente").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.TipoDerechohabiente).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Situacion Afiliatoria").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.SituacionAfiliatoria).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Datos Laborales").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.personalInfo.SituacionAfiliatoria).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                r++
            
                if(x.datosLaborales.plazasTrabajador){
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Plazas Trabajador").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    x.datosLaborales.plazasTrabajador.forEach(el => {
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Ramo").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.ramo).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Pagaduría").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.pagaduria).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Información Laboral").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                el.jobInfo.forEach(el2 =>{
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Nombramiento").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el2.Nombramiento).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Fecha de alta").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el2.fechaDeAlta).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Sueldo Básico").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el2.sueldoBasico).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Remuneración Total").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"})
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el2.remuneracionTotal).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Modalidad").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el2.Modalidad).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                })
                })
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Plazas Pensionado").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.datosLaborales.plazasPensionado).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                }
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Domicilio").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Calle").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.domicilio.Calle).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Número exterior").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.domicilio.Calle).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Número interior").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.domicilio.numeroInterior).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Colonia").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.domicilio.Colonia).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Delegación o municipio").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.domicilio.DelegacionOMunicipio).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Estado").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.domicilio.Estado).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Código postal").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.domicilio.codigoPostal).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Clínica").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Código postal").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.clinica.clinica).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Domicilio").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.clinica.Domicilio).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Colonia").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.clinica.Colonia).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Delegación ISSSTE").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.clinica.delegacionISSSTE).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Teléfono").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.clinica.telefono).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                if(x.historialCotizacion){
                    r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Historial Cotización").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    x.historialCotizacion.forEach(el => {
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Ramo").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.ramo).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Pagaduría").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.pagaduria).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    r++
                    if(el.jobInfo){
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Información Laboral").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    r++
                    r++
                    el.jobInfo.forEach(el2 => {
                        workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Tipo").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el2.Tipo).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    r++
                        workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Fecha de inicio").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el2.fechaDeInicio).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    r++
                        workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Fecha de término").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el2.fechaDeTermino).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    r++
                        workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Cotiza").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el2.Cotiza).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    r++
                        workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Sueldo Basico").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el2.sueldoBasico).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    r++
                    r++
                    })
                    }
                    
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Pagaduría").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(el.pagaduria).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                    r++
                    })
                
                }
                r++
                    workbook.sheet("Reporte de ingresos").cell(`B${r}`).value("Regimen Pensionario").style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                workbook.sheet("Reporte de ingresos").cell(`C${r}`).value(x.regimenPensionario.antiguedadParaPensiones).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
                
                for (let i = startRow; i <= r; i++) {
                    if (i % 2 === 0) {
                      workbook.sheet("Reporte de ingresos").range(`B${i}:C${i}`).style({
                        "fill": {
                          type: "pattern",
                          pattern: "solid",
                          foreground: {
                            rgb: "ffffff"
                          },
                          background: {
                            theme: 3,
                            tint: 0.4
                          }
                        }
                      });
                    } else {
                      workbook.sheet("Reporte de ingresos").range(`B${i}:C${i}`).style({
                        "fill": {
                          type: "pattern",
                          pattern: "solid",
                          foreground: {
                            rgb: "F0F0F0"
                          },
                          background: {
                            theme: 3,
                            tint: 0.4
                          }
                        }
                      });
                    }
                }
            } else{
                // "CURP inválido o no encontrado en el ISSSTE"
                workbook.sheet("Reporte de ingresos").cell(`B${r}`).value(data.issste?.data?.message).style({fontSize:11, verticalAlignment:"center", horizontalAlignment:"left", fontColor:"525151"});
                r++
            }
            
         workbook.toFileAsync("./output.xlsx")
            .then( () => {
                // Send the file as a download to the client
                res.download("./output.xlsx");
            })
        });
    } catch (err) {
        console.error(err);
        res.status(500).send(err);
    }

});
        