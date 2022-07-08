/* eslint-disable no-undef */
// eslint-disable-next-line no-undef
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Running Excel");
    $(document).ready(CargarDatos);
  }
});

// Objeto plan de cuentas
const plandecuentas = {
  codigoCuenta: "",
  nombreCuenta: "",
  descripcionCuenta: "",
};

var ObjetoPC = {};

//Creación de los controles de la paguina
$(function () {
  $("#form")
    .dxForm({
      formData: plandecuentas,
      labelMode: "floating",
      items: [
        {
          dataField: "codigoCuenta",
          label: {
            text: "Código de la Cuenta",
          },
          editorOptions: {
            showClearButton: true,
            mask: "0.0.00.00/000",
            maskChar: "-",
          },
          validationRules: [
            {
              type: "required",
              message: "Código de la Cuenta es Requerido.",
            },
            {
              type: "async",
              message: "Código de la Cuenta ya existe.",
              validationCallback(params) {
                return enviarRespuesta(params.value);
              },
            },
          ],
        },
        {
          dataField: "nombreCuenta",
          isRequired: "true",
          label: {
            text: "Nombre de la Cuenta",
          },
          editorOptions: {
            showClearButton: true,
          },
        },
        {
          dataField: "descripcionCuenta",
          editorType: "dxTextArea",
          label: {
            text: "Descripción de la Cuenta",
          },
          editorOptions: {
            showClearButton: true,
          },
        },
        {
          itemType: "button",
          buttonOptions: {
            text: "Guardar",
            stylingMode: "outlined",
            type: "default",
            useSubmitBehavior: true,
          },
        },
      ],
    })
    .dxForm("instance");
  //Funcion para guardar los datos en la hoja de claculo
  $("#form-container").on("submit", async function (e) {
    e.preventDefault();
    try {
      await Excel.run(async (context) => {
        let sheetCuentas = context.workbook.worksheets.getItemOrNullObject("Plan de Cuentas");
        await context.sync();
        // eslint-disable-next-line office-addins/load-object-before-read
        if (sheetCuentas.isNullObject) {
          sheetCuentas = context.workbook.worksheets.add("Plan de Cuentas");
          //Creat la tabla.
          let tablaCuentas = sheetCuentas.tables.add("A1:C1", true);
          //Establecer el nombre de la tabla.
          tablaCuentas.name = "PlanCuentas";
          //Colocar las cabeceras de la tabla.
          tablaCuentas.getHeaderRowRange().values = [["Código", "Cuenta", "Descripción"]];
          //tablaCuentas.columns.getItemAt(0).getDataBodyRange().numberFormat = [["#-#-##-##-###"]];
          sheetCuentas.getUsedRange().format.autofitColumns();
          sheetCuentas.getUsedRange().format.autofitRows();
          //Notificación.
          DevExpress.ui.notify({
            message: "Se CREARON la hoja Plan de Cuentas y la Tabla. Por Favor Vuelva a Guardar los datos.",
            width: 230,
            type: "warning",
            displayTime: 4000,
            animation: {
              show: {
                type: "fade",
                duration: 400,
                from: 0,
                to: 1,
              },
              hide: { type: "fade", duration: 40, to: 0 },
            },
          });
        } else {
          //Obtiene la Tabla PlanCuentas.
          let tablaCuentas = sheetCuentas.tables.getItem("PlanCuentas");
          await context.sync();
          //Agrega los datos de los textbox a la variable cuenta.
          let cuenta = [plandecuentas.codigoCuenta, plandecuentas.nombreCuenta, plandecuentas.descripcionCuenta];
          console.log(cuenta);
          //Agrega los datos a la hoja de calculo en la tabla planCuentas.
          tablaCuentas.rows.add(null, [cuenta], true);
          //Ajusta las columnas de la tabla.
          sheetCuentas.getUsedRange().format.autofitColumns();
          sheetCuentas.getUsedRange().format.autofitRows();
          //Notificacion
          DevExpress.ui.notify({
            message: `Se creó la Cuenta: ${plandecuentas.nombreCuenta}`,
            width: 230,
            type: "success",
            displayTime: 1000,
            position: "top center",
            animation: {
              show: {
                type: "fade",
                duration: 200,
                from: 0,
                to: 1,
              },
              hide: { type: "fade", duration: 20, to: 0 },
            },
          });
        }
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  });
});

// Función para carga los datos de la hoja en la constante plandecuentas
async function CargarDatos() {
  await Excel.run(async (context) => {
    let sheetCuentas = context.workbook.worksheets.getItem("Plan de Cuentas");
    let tableCuentas = sheetCuentas.tables.getItem("PlanCuentas");
    let bodyRange = tableCuentas.getDataBodyRange().load("values");

    await context.sync();
    let bodyValues = bodyRange.values;
    await context.sync();

    var reforma = bodyValues.map((value) => {
      //console.log(value);
      var obj = {};
      obj.codigoCuenta = value[0];
      obj.nombreCuenta = value[1];
      obj.descripcionCuenta = value[2];

      return obj;
    });
    ObjetoPC = reforma;
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

//Funcion
const enviarRespuesta = function (value) {
  console.log(value);
  const codigo = ObjetoPC.findIndex((obj) => obj.codigoCuenta == value);
  console.log(codigo);
  const d = $.Deferred();
  setTimeout(() => {
    d.resolve(codigo === -1);
  }, 1000);
  return d.promise();
};
