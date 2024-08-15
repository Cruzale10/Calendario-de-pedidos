document.addEventListener('DOMContentLoaded', function() {
    var calendarEl = document.getElementById('calendar');
    var loadFileButton = document.getElementById('loadFile');
    var eventDetails = document.getElementById('eventDetails');
    var eventTitle = document.getElementById('eventTitle');
    var eventDescription = document.getElementById('eventDescription');
    var closeDetailsButton = document.getElementById('closeDetails');

    var calendar = new FullCalendar.Calendar(calendarEl, {
        initialView: 'dayGridMonth',
        locale: 'es',
        events: [],
        eventClick: function(info) {
            eventTitle.textContent = info.event.title;
            eventDescription.textContent = info.event.extendedProps.description;
            eventDetails.style.display = 'block';
            info.jsEvent.preventDefault(); // Evita la acción predeterminada
        },
        eventClassNames: function(arg) {
            // Asignar clase basada en el estado
            return [arg.event.extendedProps.status];
        },
        contentHeight: 'auto' // Ajustar la altura del contenido para permitir que se muestre completo
    });

    function parseExcelDate(serial) {
        var utcDays = serial - 25569; // Ajuste para el origen de la fecha en Excel
        var date = new Date(utcDays * 86400 * 1000); // Convertir a milisegundos
        var dateString = date.toISOString().split('T')[0]; // Formato yyyy-mm-dd
        return dateString;
    }

    function getColor(status) {
        if (typeof status !== 'string') {
            console.warn('Estado no es una cadena de texto o es undefined:', status);
            return 'grey'; // Color de fondo por defecto
        }

        switch(status.trim().toLowerCase()) {
            case 'bien': return 'bien'; // Clase para estado "bien"
            case 'mal': return 'mal'; // Clase para estado "mal"
            case 'en tiempo': return 'en tiempo'; // Clase para estado "en tiempo"
            default: return 'grey'; // Color de fondo por defecto
        }
    }

    function groupBy(array, key) {
        return array.reduce((result, currentValue) => {
            const groupKey = currentValue[key];
            if (!result[groupKey]) {
                result[groupKey] = [];
            }
            result[groupKey].push(currentValue);
            return result;
        }, {});
    }

    function loadExcelData() {
        var input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.csv';

        input.onchange = function(event) {
            var file = event.target.files[0];
            var reader = new FileReader();
            
            reader.onload = function(e) {
                var data = new Uint8Array(e.target.result);
                var workbook = XLSX.read(data, {type: 'array'});
                var firstSheetName = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[firstSheetName];
                var jsonData = XLSX.utils.sheet_to_json(worksheet);

                // Agrupar los datos por "OVTA"
                var groupedData = groupBy(jsonData, 'OVTA');
                
                var events = Object.keys(groupedData).map(function(ovta) {
                    var group = groupedData[ovta];
                    
                    return {
                        title: 'Cliente: ' + group[0].Cliente + ' - OVTA: ' + ovta,
                        start: parseExcelDate(group[0].Fecha), // Usar la fecha del primer pedido en el grupo
                        backgroundColor: getColor(group[0]['Estatus']), // Color basado en el estatus de la columna M
                        extendedProps: {
                            description: group.map(function(row) {
                                return 'Código: ' + row.Codigo + 
                                       '\nMonto Neto: ' + row['Monto neto'] + 
                                       '\nProducto: ' + row['Nombre del producto'];
                            }).join('\n\n'), // Combina la información de todos los pedidos del grupo
                            status: getColor(group[0]['Estatus']) // Asignar la clase del estado
                        }
                    };
                });

                calendar.addEventSource(events);
            };
            
            reader.readAsArrayBuffer(file);
        };

        input.click();
    }

    loadFileButton.addEventListener('click', loadExcelData);

    closeDetailsButton.addEventListener('click', function() {
        eventDetails.style.display = 'none';
    });

    calendar.render();
});
