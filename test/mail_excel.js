document.getElementById('dataForm').addEventListener('submit', function(event) {
    event.preventDefault();

    const formData = new FormData(event.target);
    const cliente = formData.get('cliente');
    const sucursal = formData.get('sucursal');
    const sector = formData.get('sector');
    const dir_origen = formData.get('dir_origen');
    const nombre_contacto = formData.get('nombre_contacto');
    const tel_contacto = formData.get('tel_contacto');
    const fecha_retiro = formData.get('fecha_retiro');
    const horario_retiro = formData.get('horario_retiro');
    const plazo = formData.get('plazo');
    const tipo_producto = formData.get('tipo_producto');
    const vehiculo = formData.get('vehiculo');
    const dir_entrega = formData.get('dir_entrega');
    const nombre_contacto_entrega = formData.get('nombre_contacto_entrega');
    const tel_contacto_entrega = formData.get('tel_contacto_entrega');
    const fecha_entrega = formData.get('fecha_entrega');
    const horario_entrega = formData.get('horario_entrega');
    const observaciones = formData.get('observaciones');
    const email = formData.get('email');

    // Crear un nuevo libro de trabajo y una hoja activa
    const wb = XLSX.utils.book_new();
    const ws_data = [
        ['FICHA PARA SOLICITUD DE CADETERÍAS.', ''],
        ['Cliente:', cliente],
        ['Sucursal:', sucursal],
        ['Sector que solicita:', sector],
        ['', ''],
        ['Dirección de Origen:', dir_origen],
        ['Nombre contacto:', nombre_contacto],
        ['Tel contacto:', tel_contacto],
        ['Fecha de retiro:', fecha_retiro],
        ['Horario de retiro:', horario_retiro],
        ['Plazo (Express, en el dia, 24hs, 48hs):', plazo],
        ['Tipo de producto a retirar:', tipo_producto],
        ['Vehículo:', vehiculo],
        ['', ''],
        ['', ''],
        ['Dirección de entrega:', dir_entrega],
        ['Nombre contacto (entrega):', nombre_contacto_entrega],
        ['Tel contacto (entrega):', tel_contacto_entrega],
        ['Fecha de entrega:', fecha_entrega],
        ['Horario de entrega:', horario_entrega],
        ['', ''],
        ['', ''],
        ['Observaciones adicionales:', observaciones]
    ];
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    XLSX.utils.book_append_sheet(wb, ws, 'Datos');

    // Generar archivo Excel
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);

    // Crear enlace de descarga
    const downloadLink = document.createElement('a');
    downloadLink.href = url;
    downloadLink.download =  'Solicitud de cadeteria '+observaciones+'.xlsx';
    downloadLink.innerText = 'Descargar archivo Excel';

    // Crear enlace mailto
    const observaciones1 = formData.get('observaciones'); // Obtenido previamente
    const email1 = "logistica@upostal.com.uy";
    const ccEmails = "expedicion@nad.uy";
    const subject = `Solicitud Express ${observaciones1}`;
    const body = `Buenos días, Estimados, espero se encuentren bien.\n\nLes realizamos una solicitud Express. Adjunto la planilla.\n\nGracias.`;
    const mailtoLink = document.createElement('a');
    mailtoLink.href = `mailto:${email1}?cc=${ccEmails}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;    mailtoLink.innerText = 'Enviar Correo';

    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = '';
    resultDiv.appendChild(downloadLink);
    resultDiv.appendChild(document.createElement('br'));
    resultDiv.appendChild(mailtoLink);
});
