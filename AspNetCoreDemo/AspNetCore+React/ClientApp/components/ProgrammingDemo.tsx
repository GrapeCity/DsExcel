import * as React from 'react';
import 'isomorphic-fetch';
import { Utility } from '../utility';
import Select from 'react-select'
import 'react-select/dist/react-select.css';

interface ProgrammingDemoState {
    options: { value: string; label: string; }[],
    value: string
}

//ServerSidePrograming
export class ProgrammingDemo extends React.Component<{}, ProgrammingDemoState> {
    spread: GC.Spread.Sheets.Workbook;

    constructor() {
        super();

        this.state = {
            options: [
                { value: 'BidTracker', label: 'Bid Tracker' },
                { value: 'ToDoList', label: 'ToDo List' },
                { value: 'AddressBook', label: 'Address Book' }
            ],
            value: 'BidTracker'
        }

        this.exportExcel = this.exportExcel.bind(this);
        this.onUseCaseChange = this.onUseCaseChange.bind(this);
    }

    public render() {
        return <div className='spread-page'>
            <h1>Programming API Demo</h1>
            <p>This example demonstrates how to program with <strong>GcExcel</strong> to generate a complete spreadsheet model at server side.  You can find all of source code in the SpreadServicesController.cs.  We use <strong>Spread.Sheets</strong> as client-side viewer. </p>
            <ul>
                <li>You can first program with <strong>GcExcel</strong> server-side.</li>
                <li><strong>GcExcel</strong> then inoke <strong>ToJson</strong> and transport the ssjson to client side.</li>
                <li>In the browser script, <strong>Spread.Sheets</strong> will invoke <strong>fromJSON</strong> with the ssjson from the server.</li>
                <li>Then, you can view the result in <strong>Spread.Sheets</strong> or download it as an excel file.</li>
            </ul>
            <br />
            <div className='btn-group'>
                <Select className='select'
                    name="form-field-name"
                    value={this.state.value}
                    options={this.state.options}
                    onChange={this.onUseCaseChange} />
                <button className='btn btn-default btn-md' onClick={this.exportExcel}>Export Excel</button>
            </div>
            <div id='spreadjs' className='spread-div' />
        </div>;
    }

    componentDidMount() {
        this.spread = new GC.Spread.Sheets.Workbook(document.getElementById('spreadjs'), {
            sheetCount: 1
        });

        this.loadSpreadFromUseCase(this.state.value);
    }

    loadSpreadFromUseCase(caseName: string) {
        var requestUrl = '/api/SpreadServices/GetSSJsonFromUseCase/' + caseName;
        fetch(requestUrl, {
            method: 'Get'
        }).then(response => response.json() as Promise<object>)
            .then(data => {
                this.spread.fromJSON(data);
            });
    }

    onUseCaseChange(newValue) {
        this.setState({ value: newValue.value });
        this.loadSpreadFromUseCase(newValue.value);
    }

    exportExcel() {
        var ssjson = JSON.stringify(this.spread.toJSON(null));
        Utility.ExportExcel(ssjson, this.state.value);
    }
}


