const express = require('express');
const xlsx = require('node-xlsx');
const fs = require('fs');
const app = express();
const port = 3000;
const dateFns = require('date-fns'); // Necesitas instalar esta librería

app.use(express.static('public'));

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

app.get('/generate-excel', (req, res) => {
  // Función para convertir un número de serie de Excel a una fecha de JavaScript
  const excelToDate = (excelDate) => {
    const secondsInDay = 24 * 60 * 60 * 1000;
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(excelEpoch.getTime() + excelDate * secondsInDay);
  };

  // Modificar la matriz de datos para convertir los números de Excel en fechas
  const data = [
    ['Nombre', 'Edad', 'Fecha de Nacimiento','dinero'],
    ['Juan',   25, excelToDate(42291), 775000],
    ['María',  30, excelToDate(42304), 450000],
    ['Carlos', 22, excelToDate(42307), 1800000]
  ];

  // Crear un libro de Excel
  const buffer = xlsx.build([{ name: 'Sheet 1', data: data }]);

  // Guardar el libro de Excel en un archivo
  const excelFileName = 'data.xlsx';
  fs.writeFileSync(excelFileName, buffer);

  // Descargar el archivo Excel
  res.download(excelFileName, () => {
    // Eliminar el archivo después de la descarga
    fs.unlinkSync(excelFileName);
  });
});

app.listen(port, () => {
  console.log(`La aplicación está en ejecución en http://localhost:${port}`);
});
