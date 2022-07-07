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
          editorType: "dxTextBox",
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
          editorType: "dxTextBox",
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
    console.log(dato.codigoCuenta);
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
        } else {
          let tablaCuentas = sheetCuentas.tables.getItem("PlanCuentas");
          await context.sync();
          let cuenta = [dato.codigoCuenta, dato.nombreCuenta, dato.descripcionCuenta];
          tablaCuentas.rows.add(null, [cuenta], true);

          dato.values = null;
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
