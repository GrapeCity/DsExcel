import * as React from 'react';
import 'isomorphic-fetch';
import { Utility } from '../utility'

//ExcelIO
export class ExcelIODemo extends React.Component<{}, {}> {
    selectedFileName: string;
    spread: GC.Spread.Sheets.Workbook;

    constructor() {
        super();

        this.selectedFileName = null;

        this.importExcel = this.importExcel.bind(this);
        this.exportExcel = this.exportExcel.bind(this);
    }

    public render() {
        return <div className='spread-page'>
            <h1>Excel Input & Output Demo</h1>
            <p>This example demonstrates how to use <strong>GcExcel</strong> as server-side spreadsheet model, and use <strong>Spread.Sheets</strong> the front-end viewer and editor.</p>
            <ul>
                <li><strong>GcExcel</strong> can import an excel file and export to ssjson format, then transport the ssjson to client-side.</li>
                <li><strong>Spread.Sheets</strong> client-side can receive and load the ssjson from server-side.</li>
                <li>You can view the content of the excel file through <strong>Spread.Sheets</strong>.</li>
                <li>You can also make changes to the content, and send the whole document as ssjson to <strong>GcExcel</strong> server-side.</li>
                <li><strong>GcExcel</strong> server-side loads the ssjson and saves to a new excel file, then you can download the modified excel file.</li>
            </ul>
            <br/>
            <div className='btn-group'>
                <input className='btn btn-default btn-md' type='file' onChange={this.importExcel} title='Import Excel' value='Choose an excel file' />
                <button className='btn btn-default btn-md' onClick={this.exportExcel}>Export Excel</button>
            </div>
            <div id='spreadjs' className='spread-div' />
        </div>;
    }

    /**
     * Upload an excel file at client side, open the file at server side then transport the ssjson to client
     * @param e
     */
    importExcel(e) {
        var selectedFile = e.target.files[0];
        if (!selectedFile) {
            this.selectedFileName = null;
            return;
        }

        this.selectedFileName = selectedFile.name;
        var requestUrl = '/api/SpreadServices/ImportExcel';
        fetch(requestUrl, {
            method: 'POST',
            body: selectedFile
        }).then(response => response.json() as Promise<object>)
            .then(data => {
                this.spread.fromJSON(data);
                //this.setState({ ssjson: data, loading: false });
            });
    }

    /**
     * Tranport ssjson from Spread.Sheets and save and download the excel file.
     * @param e
     */
    exportExcel(e) {
        var ssjson = JSON.stringify(this.spread.toJSON(null));
        Utility.ExportExcel(ssjson, this.selectedFileName);
    }

    componentDidMount() {
        this.spread = new GC.Spread.Sheets.Workbook(document.getElementById('spreadjs'), {
            sheetCount: 1
        });

       var sheet = this.spread.getActiveSheet();
        sheet.addSpan(6, 3, 5, 10);
        var cell = sheet.getCell(6, 3);
        cell.text('Please choose an excel file(xlsx) to import!');
        cell.hAlign(GC.Spread.Sheets.HorizontalAlign.center);
        cell.vAlign(GC.Spread.Sheets.VerticalAlign.center);
        cell.font('bold 25px arial');
    }
}


