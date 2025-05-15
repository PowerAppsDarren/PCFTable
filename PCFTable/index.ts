import { IInputs, IOutputs } from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;

export class PCFTable implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;

    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        const messageElement = document.createElement("div");
        messageElement.id = "recordCountMessage"; // Good practice to give elements IDs
        // Add some very basic inline styling (not best practice for larger components, but ok for intro)
        messageElement.style.padding = "10px";
        messageElement.style.fontSize = "16px";
        messageElement.innerText = "Initializing..."; // Initial text
        this._container.appendChild(messageElement); // Add it to the component's main container

        // Ensure the container is empty if we are re-initializing (not typical, but good for iterative development)
        while (container.firstChild) {
            container.removeChild(container.firstChild);
        }
        const tableContainer = document.createElement("div");
        tableContainer.id = "customTableContainer";
        this._container.appendChild(tableContainer);
    }

    /**
     * Called when any value in the property bag has changed. 
     * This includes field values, data-sets, global values 
     * such as container height and width, offline status, 
     * control metadata values such as label, visible, etc.
     * 
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {

        const dataSet = context.parameters.sampleDataSet; // 'sampleDataSet' must match the name in ControlManifest.Input.xml
        const messageElement = this._container.querySelector("#recordCountMessage") as HTMLDivElement; // Get the element we created in init

        if (dataSet.loading) { // Check if data is still loading
            messageElement.innerText = "Loading data...";
            return;
        }

        const recordCount = dataSet.sortedRecordIds.length;
        messageElement.innerText = `Number of records: ${recordCount}`;

        // Example: Log dataset columns to console for inspection (useful for debugging)
        if (dataSet.columns && dataSet.columns.length > 0) {
            console.log("Dataset Columns:", dataSet.columns.map(col => col.displayName));
        }


        //const dataSet = context.parameters.sampleDataSet;
        const tableContainer = this._container.querySelector("#customTableContainer") as HTMLDivElement;

        // Clear previous table content
        while (tableContainer.firstChild) {
            tableContainer.removeChild(tableContainer.firstChild);
        }

        if (dataSet.loading) {
            tableContainer.innerText = "Loading data...";
            return;
        }

        if (dataSet.sortedRecordIds.length === 0) {
            tableContainer.innerText = "No records to display.";
            return;
        }

        // Create table element
        const table = document.createElement("table");
        table.style.borderCollapse = "collapse"; // Basic styling
        table.style.width = "100%";

        // Create table header
        const a_column_name = "ID"; // Replace this with a valid column name from your intended data source
        const another_column_name = "Joke Setup"; // Replace this as well

        // Get actual display names (more robust)
        // For simplicity now, let's assume we know a couple of column *logical names* we want to display
        // Students will need to know or find these logical names from their Power Apps data source (e.g., 'cr47f_productname', 'emailaddress1')
        // A more dynamic approach would iterate through dataSet.columns, but let's start simpler.
        const columnDefinitions = [
            { logicalName: a_column_name, displayName: "Column 1 Display" }, // User should replace these
            { logicalName: another_column_name, displayName: "Column 2 Display" }  // User should replace these
        ];
        // IMPORTANT: Remind students they'll need to update 'a_column_name' and 'another_column_name'
        // with actual logical names from the data source they intend to use in Power Apps.
        // How to find logical names: In Power Apps, go to the table settings, columns, and look for the 'Name'.

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        columnDefinitions.forEach(colDef => {
            const th = document.createElement("th");
            th.innerText = colDef.displayName; // Use predefined display names for now
            th.style.border = "1px solid black";
            th.style.padding = "5px";
            th.style.textAlign = "left";
            headerRow.appendChild(th);
        });

        // Create table body and rows
        const tbody = table.createTBody();
        dataSet.sortedRecordIds.forEach(recordId => {
            const record = dataSet.records[recordId];
            const row = tbody.insertRow();
            columnDefinitions.forEach(colDef => {
                const cell = row.insertCell();
                cell.innerText = record.getFormattedValue(colDef.logicalName) || ""; // Use logical name to fetch
                cell.style.border = "1px solid black";
                cell.style.padding = "5px";
            });
        });

        tableContainer.appendChild(table);

        
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, 
     * expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
