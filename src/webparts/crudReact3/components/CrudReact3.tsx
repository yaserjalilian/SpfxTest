import * as React from 'react';
import styles  from './CrudReact3.module.scss'
import { ICrudReact3Props } from './ICrudReact3Props';
import { escape } from '@microsoft/sp-lodash-subset';
import { Version } from '@microsoft/sp-core-library';
import { DatePicker, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { sp, Web, IWeb, List} from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { BaseWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib';
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import CrudReact3WebPart, { ICrudReact3WebPartProps } from '../CrudReact3WebPart';
import { Link } from 'office-ui-fabric-react';
import { FontIcon } from '@fluentui/react/lib/Icon';
export interface IStates {
  Items: any;
  ID: any;
  Owner: any;
  OwnerId: any;
  HireDate: any;
  Destination: any;
  OrderNumber: any;
  CustomerName: any;
  State: any;
  HTML: any;
  LinkToFile: any;
}

const fieldUpdateValues = { 
  Tags: ["Pending", "Retreived"],
  "Tags@odata.type": "Collection(Edm.String)"
};

export default class CrudReact3 extends React.Component<ICrudReact3Props, IStates> {
  constructor(props) {
      super(props);
      this.state = {
        Items: [],
        Owner: "",
        OwnerId: 0,
        ID: 0,
        HireDate: null,
        Destination: "",
        OrderNumber: 0,
        CustomerName: " ",
        State: "Pending",
        HTML: [],
        LinkToFile:""
  
      };

    }
    
    public async componentDidMount() {
      await this.fetchData();
    }
   
    public  async  saveIntoSharePoint(file: IFilePickerResult) {
      let siteUrl = this.props.webURL;
      let web =  Web(siteUrl);
      if (file.fileAbsoluteUrl == null) {
        file.downloadFileContent().then(async r => {
        if (r.size <= 10485760) {
          // small upload
          var fileUploaded=   web.getFolderByServerRelativeUrl("/Shared%20Documents/").files.add(file.fileName, r, true);
        }
        
        else {

          // large upload
        var fileUploaded=  web.getFolderByServerRelativeUrl("/Shared%20Documents/").files.addChunked(file.fileName, r,data => {}, true)
      }
      
          });

      }
      else {

      }
      this.setState({LinkToFile:siteUrl+"/Shared%20Documents/"+file.fileName})
       }
    public async fetchData() {
       
      var web = Web(this.props.webURL);
      const items: any[] = await web.lists.getByTitle("Orders").items.select("*", "Owner/Title").expand("Owner/ID").get();
      console.log(items);
      this.setState({ Items: items });
      let html = await this.getHTML(items);
      this.setState({ HTML: html });
    }
    public findData = (id): void => {
      this.fetchData();
      var itemID = id;
      var allitems = this.state.Items;
      var allitemsLength = allitems.length;
      if (allitemsLength > 0) {
        for (var i = 0; i < allitemsLength; i++) {
          if (itemID == allitems[i].Id) {
            this.setState({
              ID:itemID,
              Owner:allitems[i].Owner.Title,
              OwnerId:allitems[i].OwnerId,
              HireDate:new Date(allitems[i].HireDate),
             Destination:allitems[i].Destination,
             OrderNumber:allitems[i].OrderNumber,
              CustomerName:allitems[i].CustomerName,
               State:allitems[i].State
             ,  LinkToFile:allitems[i].LinkToFile
            });
        }
      }
    }
  
    }
   public async getHTML(items) {
        var tabledata = <table className={styles.table}>
          <thead>
            <tr>
              <th>Order Number</th>
              <th>Customer Name</th>
              <th>Destination</th>
              <th>Owner</th>
              <th>State</th>
              <th>Link to File</th>
            </tr>
          </thead>
          <tbody>
            {items && items.map((item, i) => {
              return [
                <tr key={i} onClick={()=>this.findData(item.ID)}>
                 <td>{item.OrderNumber}</td> 
                 <td>{item.CustomerName}</td>
                 <td>{item.Destination}</td>
                  <td>{item.Owner.Title}</td>
                  <td>{item.State}</td>
                  <td>{FormatDate(item.HireDate)}</td>
                 
               <td> <Link href={item.LinkToFile} target='_blank'>  <FontIcon iconName="Dictionary"  /> </Link></td>
                </tr>
              ];
            })}
          </tbody>
    
        </table>;
        return await tabledata;
      }
      public _getPeoplePickerItems = async (items: any[]) => {
    
        if (items.length > 0) {
    
          this.setState({ Owner: items[0].text });
          this.setState({ OwnerId: items[0].id });
        }
        else {
          //ID=0;
          this.setState({ OwnerId: "" });
          this.setState({ Owner: "" });
        }
      }
      public onchange=(e,stateValue)=> {
        var state = {};
        state[stateValue] = e.target.value;
        this.setState(state);
        
      }

      public setstatelocal=( x)=> {
        var state = {};
        state["LinkToFile"] = x;
        this.setState(state);
        
      }
      
      private async SaveData() {
        let web = Web(this.props.webURL);
        
        

        await web.lists.getByTitle("Orders").items.add({
    
          OwnerId:this.state.OwnerId,
          HireDate: new Date(this.state.HireDate),
         Destination: this.state.Destination,
         OrderNumber: this.state.OrderNumber, 
         CustomerName: this.state.CustomerName, 
         State: this.state.State, 
          LinkToFile: this.state.LinkToFile
        }). then(i => {
          console.log(i);
        });
        alert("Created Successfully");
        this.setState({Owner:"",HireDate:null,Destination:"",OrderNumber:"",CustomerName:"",State:"",LinkToFile:""});
        this.fetchData();
      }
      private async UpdateData() {
        let web = Web(this.props.webURL);
        await web.lists.getByTitle("Orders").items.getById(this.state.ID).update({
          OwnerId:this.state.OwnerId,
          OrderNumber: this.state.OrderNumber,
          State: this.state.State,
          CustomerName: this.state.CustomerName,
          HireDate: new Date(this.state.HireDate),
          Destination: this.state.Destination,
     
        }).then(i => {
          console.log(i);
        });
        alert("Updated Successfully");
        this.setState({Owner:"",HireDate:null,Destination:"",OrderNumber:"",State:"",CustomerName:"",LinkToFile:""});
        this.fetchData();
      }
      private async DeleteData() {
        let web = Web(this.props.webURL);
        await web.lists.getByTitle("Orders").items.getById(this.state.ID).delete()
        .then(i => {
          console.log(i);
        });
        alert("Deleted Successfully");
        this.setState({Owner:"",HireDate:null,Destination:"",OrderNumber:"",State:"",CustomerName:"",LinkToFile:""});
        this.fetchData();
      }
      public render(): React.ReactElement<ICrudReact3WebPartProps> {
        return (
          <div>
            <h1>CRUD Operations With ReactJs</h1>
            {this.state.HTML}
            <div className={styles.btngroup}>
              <div><PrimaryButton text="Create" onClick={() => this.SaveData()}/></div>
              <div><PrimaryButton text="Update" onClick={() => this.UpdateData()} /></div>
              <div><PrimaryButton text="Delete" onClick={() => this.DeleteData()}/></div>
            </div>
            <div>
              <form>

              <div>
                  <Label>Order Number</Label>
                  <TextField defaultValue=' ' value={this.state.OrderNumber}  onChange={(value) => this.onchange(value, "OrderNumber")} />
                </div>

                <div>
                  <Label>Customer Name</Label>
                  <TextField defaultValue=' ' value={this.state.CustomerName}  onChange={(value) => this.onchange(value, "CustomerName")} />
                </div>

                <div>
                  <Label>Destination</Label>
                  <TextField defaultValue=' ' value={this.state.Destination} multiline onChange={(value) => this.onchange(value, "Destination")} />
                </div>

                {/* <div>
                  <Label>State</Label>
                  <TextField defaultValue=' ' value={this.state.State}  onChange={(value) => this.onchange(value, "OrderNumberState")} />
                </div> */}

                <div>
                  <Label>Owner</Label>
                  <PeoplePicker
                                        context={this.props.context as any}
                                        personSelectionLimit={1}
                                        // defaultSelectedUsers={this.state.Owner===""?[]:this.state.Owner}
                                        isRequired={false}
                                        defaultSelectedUsers={[this.state.Owner?this.state.Owner:""]}
                                        showHiddenInUI={false}
                                        principalTypes={[PrincipalType.User]}
                                        resolveDelay={1000}
                                        ensureUser={true}
                                        selectedItems={this._getPeoplePickerItems}                                        
                                        />
                </div>
                  <div>
                  <Label>Date</Label>
                  <DatePicker maxDate={new Date()} allowTextInput={false} strings={DatePickerStrings} value={this.state.HireDate} onSelectDate={(e) => { this.setState({ HireDate: e }); }} ariaLabel="Select a date" formatDate={FormatDate} />
                </div>
                
                <div >
        {/* <img src={this.state.LinkToFile} height={'150px'} width={'150px'}></img> */}
        <br></br>
        <br></br>
        <FilePicker
          label={'Select or upload file'}
          buttonClassName={styles.button}
          buttonLabel={'Images'}
          accepts={[".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"]}
          buttonIcon="FileImage"
          onSave={this.saveIntoSharePoint.bind(this)}
          onChanged={this.saveIntoSharePoint.bind(this)}
        
          context={this.props.context}
        />
      </div>
    {/* onSave={(filePickerResult: IFilePickerResult) => { this.setState({ LinkToFile: filePickerResult.fileAbsoluteUrl });  alert(JSON.stringify(this.state.LinkToFile)); }}    
          onChanged={(filePickerResult: IFilePickerResult) => { this.setState({LinkToFile: filePickerResult.fileAbsoluteUrl} );alert(JSON.stringify(this.state.LinkToFile)); }}     */}
              </form>
            </div>
          </div>
        );
      }
  }
  export const DatePickerStrings: IDatePickerStrings = {
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
    goToToday: 'Go to today',
    prevMonthAriaLabel: 'Go to previous month',
    nextMonthAriaLabel: 'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year',
    invalidInputErrorMessage: 'Invalid date format.'
  };
  export const FormatDate = (date): string => {
    console.log(date);
    var date1 = new Date(date);
    var year = date1.getFullYear();
    var month = (1 + date1.getMonth()).toString();
    month = month.length > 1 ? month : '0' + month;
    var day = date1.getDate().toString();
    day = day.length > 1 ? day : '0' + day;
    return month + '/' + day + '/' + year;
  };
