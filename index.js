const express = require('express');
const app = express();
const excel = require('exceljs');
const fs = require('fs');

let workbook;

app.use(express.urlencoded({ extended: false }));
app.use(express.static('public'));

function crearOcargarExcel() {
  workbook = new excel.Workbook();

  const existeArchivo = fs.existsSync('datos.xlsx');

  if (existeArchivo) {
    workbook.xlsx.readFile('datos.xlsx')
      .catch(err => {
        console.error('Error al cargar el archivo Excel:', err);
      });
  } else {
    const worksheet = workbook.addWorksheet('Datos');
    worksheet.addRow(['Nombre', 'Email', 'Teléfono']);
  }
}

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

app.post('/formulario', (req, res) => {
  const { nombre, email, telefono } = req.body;

  if (!workbook) {
    crearOcargarExcel();
  }

  const worksheet = workbook.getWorksheet('Datos');
  worksheet.addRow([nombre, email, telefono]);

  workbook.xlsx.writeFile('datos.xlsx')
    .then(() => {
      console.log('Información agregada al archivo Excel exitosamente');
      res.send('¡Formulario enviado con éxito!');
    })
    .catch(err => {
      console.error('Error al agregar información al archivo Excel:', err);
      res.status(500).send('Ocurrió un error al procesar el formulario');
    });
});

const port = 3000;
app.listen(port, () => {
  console.log(`Servidor web iniciado en http://localhost:${port}`);
});

