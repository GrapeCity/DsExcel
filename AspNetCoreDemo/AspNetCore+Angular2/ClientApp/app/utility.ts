import 'isomorphic-fetch';

export class Utility {

    public static ExportExcel(ssjon: string, fileName: string) {
        var requestUrl = '/api/SpreadServices/ExportExcel';
        fetch(requestUrl, {
            method: 'POST',
            body: ssjon
        }).then(function (response) {
            var blob = response.blob();
            return blob;
        }).then(blob => {
            if (!fileName) {
                fileName = 'Spread.Services-exported.xlsx';
            }
            Utility.DownloadFile(blob, fileName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        });
    }

    public static DownloadFile(data: Blob, filename: string, mime: string): void {
        var blob = new Blob([data], { type: mime || 'application/octet-stream' });
        if (typeof window.navigator.msSaveBlob !== 'undefined') {
            // IE workaround for "HTML7007: One or more blob URLs were 
            // revoked by closing the blob for which they were created. 
            // These URLs will no longer resolve as the data backing 
            // the URL has been freed."
            window.navigator.msSaveBlob(blob, filename);
        }
        else {
            var blobURL = window.URL.createObjectURL(blob);
            var tempLink = document.createElement('a');
            tempLink.href = blobURL;
            tempLink.setAttribute('download', filename);
            tempLink.setAttribute('target', '_blank');
            document.body.appendChild(tempLink);
            tempLink.click();
            document.body.removeChild(tempLink);
        }
    }
}