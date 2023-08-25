const express = require('express');
const xlsx = require('node-xlsx');
const fs = require('fs');
const app = express();
const port = 3000;

app.use(express.static('public'));

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

app.get('/generate-excel', (req, res) => {
  // Función para convertir un número de serie de Excel a una fecha de JavaScript
  // const excelToDate = (excelDate) => {
  //   const secondsInDay = 24 * 60 * 60 * 1000;
  //   const excelEpoch = new Date(Date.UTC(1900,0,0));
  //   return new Date(excelEpoch.getTime() + excelDate * secondsInDay);
  // };

  //La misma función pero con retorno de string
  const excelToDate = (excelDate) => {
    const secondsInDay = 24 * 60 * 60 * 1000;
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const dateWithTime = new Date(excelEpoch.getTime() + excelDate * secondsInDay);

    // Si la fecha es anterior al 1 de marzo de 1900, ajustarla restando un día
    if (excelDate < 61) {
        dateWithTime.setUTCDate(dateWithTime.getUTCDate() - 1);
    }

    // Formatear la fecha como "DD/MM/YYYY"
    const year = dateWithTime.getUTCFullYear();
    const month = (dateWithTime.getUTCMonth() + 1).toString().padStart(2, '0');
    const day = dateWithTime.getUTCDate().toString().padStart(2, '0');
    
    return `${day}/${month}/${year}`;
};

  
  

  // Modificar la matriz de datos para convertir los números de Excel en fechas
  const data = [
    ['Nombre', 'Edad', 'Fecha de Nacimiento', 'Dinero','Porcentaje'],
    ['Juan', 25, 42291, 775000,  0.2400845],
    ['María', 30, 42304, 450000,  0.23645],
    ['Carlos', 22, 42307, 1800000,  0.1935]
  ];

  // Crear un libro de Excel
  const buffer = xlsx.build([{ name: 'Sheet 1', data: data }]);

  // Obtener la hoja de trabajo
  const ws = xlsx.parse(buffer)[0].data;

  // Aplicar formato de porcentaje a la columna "Porcentaje" (columna 4)
  for (let i = 1; i < ws.length; i++) {
    if (ws[i][4] != null) {
      ws[i][4] = { t: 'n', z: '0.00%', v: ws[i][4] };
    }
  }

  // Aplicar formato de moneda a la columna "Dinero" (columna 3)
  for (let i = 1; i < ws.length; i++) {
    if (ws[i][3] != null) {
      ws[i][3] = { t: 'n', z: '"$"#,##0.00', v: ws[i][3] };
    }
  }

  // Aplicar formato de fecha a la columna "Fecha de Nacimiento" (columna 2)
  for (let i = 1; i < ws.length; i++) {
    if (ws[i][2] != null) {
      console.log(ws[i][2]);
      console.log(excelToDate(ws[i][2]));
      //Formato string
      ws[i][2] = { t: 's',z: 'dd/mm/yyyy', v: excelToDate(ws[i][2]) };
      //Formato date pero con hora incluida
      // ws[i][2] = { t: 'd',z: 'dd/mm/yyyy', v: excelToDate(ws[i][2]) };

    }
  }

  // Crear un nuevo libro de Excel con los datos formateados
  const formattedBuffer = xlsx.build([{ name: 'Sheet 1', data: ws }]);

  // Guardar el libro de Excel en un archivo
  const excelFileName = 'data.xlsx';
  fs.writeFileSync(excelFileName, formattedBuffer);

  // Descargar el archivo Excel
  res.download(excelFileName, () => {
    // Eliminar el archivo después de la descarga
    fs.unlinkSync(excelFileName);
  });
});

app.listen(port, () => {
  console.log(`La aplicación está en ejecución en http://localhost:${port}`);
});
