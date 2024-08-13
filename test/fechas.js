// Obtener la fecha de hoy
var today = new Date();
var day = String(today.getDate()).padStart(2, '0');
var month = String(today.getMonth() + 1).padStart(2, '0'); // Enero es 0
var year = today.getFullYear();

// Formatear la fecha como YYYY-MM-DD
var todayFormatted = year + '-' + month + '-' + day;

// Establecer el valor del campo de fecha
document.getElementById('fecha_retiro').value = todayFormatted;
document.getElementById('fecha_entrega').value = todayFormatted;