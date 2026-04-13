async function RequestExportExcel(dataExcel) {
    if (!dataExcel || typeof dataExcel !== 'object') {
        throw new Error('dataExcel debe ser un objeto válido');
    }

    const endpoint = window.location.origin + '/exportExcel.php';

    try {
        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(dataExcel)
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error('Error servidor:', response.status, errorText);
            return;
        }

        let fileName = 'reporte.xlsx';
        const disposition = response.headers.get('Content-Disposition');

        if (disposition) {
            const match = disposition.match(/filename="?([^"]+)"?/i);
            if (match && match[1]) fileName = match[1];
        }

        const blob = await response.blob();
        const blobUrl = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        a.remove();

        URL.revokeObjectURL(blobUrl);

    } catch (error) {
        console.error('Error:', error);
    }
}

window.chartColors = {
    red: 'rgb(255, 99, 132)',
    orange: 'rgb(255, 159, 64)',
    yellow: 'rgb(255, 205, 86)',
    green: 'rgb(75, 192, 192)',
    blue: 'rgb(54, 162, 235)',
    purple: 'rgb(153, 102, 255)',
    grey: 'rgb(231,233,237)'
};

window.randomScalingFactor = function () {
    return (Math.random() > 0.5 ? 1.0 : -1.0) * Math.round(Math.random() * 100);
}

