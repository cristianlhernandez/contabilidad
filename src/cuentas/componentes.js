/* eslint-disable no-undef */
// eslint-disable-next-line no-undef
$(function () {
  $("#form")
    .dxForm({
      formData: dato,
      labelMode: "floating",
      items: [
        {
          dataField: "codigoCuenta",
          isRequired: "true",
          label: {
            text: "Código de la Cuenta",
          },
          editorOptions: {
            showClearButton: true,
            mask: "0.0.00.00/000",
            maskChar: "-",
          },
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
          let tablaCuentas = sheetCuentas.tables.getItem("PlanCuentas");
          await context.sync();
          let cuenta = [dato.codigoCuenta, dato.nombreCuenta, dato.descripcionCuenta];
          tablaCuentas.rows.add(null, [cuenta], true);
          sheetCuentas.getUsedRange().format.autofitColumns();
          sheetCuentas.getUsedRange().format.autofitRows();
          DevExpress.ui.notify({
            message: `Se creó la Cuenta: ${dato.nombreCuenta}`,
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

const dato = [
  {
    codigoCuenta: "",
    nombreCuenta: "",
    descripcionCuenta: "",
  },
];
