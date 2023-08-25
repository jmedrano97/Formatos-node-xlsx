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
  // Crear un objeto de datos para el archivo Excel
  const data = [
    ['Nombre', 'Edad', 'Fecha de Nacimiento'],
    ['Juan',   25, 42291  ],
    ['María',  30, 42304  ],
    ['Carlos', 22, 42307  ]
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
