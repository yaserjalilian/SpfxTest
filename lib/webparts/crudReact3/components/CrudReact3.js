var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import styles from './CrudReact3.module.scss';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Web } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { FilePicker } from '@pnp/spfx-controls-react/lib';
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { Link } from 'office-ui-fabric-react';
import { FontIcon } from '@fluentui/react/lib/Icon';
var fieldUpdateValues = {
    Tags: ["Pending", "Retreived"],
    "Tags@odata.type": "Collection(Edm.String)"
};
var CrudReact3 = /** @class */ (function (_super) {
    __extends(CrudReact3, _super);
    function CrudReact3(props) {
        var _this = _super.call(this, props) || this;
        _this.findData = function (id) {
            _this.fetchData();
            var itemID = id;
            var allitems = _this.state.Items;
            var allitemsLength = allitems.length;
            if (allitemsLength > 0) {
                for (var i = 0; i < allitemsLength; i++) {
                    if (itemID == allitems[i].Id) {
                        _this.setState({
                            ID: itemID,
                            Owner: allitems[i].Owner.Title,
                            OwnerId: allitems[i].OwnerId,
                            HireDate: new Date(allitems[i].HireDate),
                            Destination: allitems[i].Destination,
                            OrderNumber: allitems[i].OrderNumber,
                            CustomerName: allitems[i].CustomerName,
                            State: allitems[i].State,
                            LinkToFile: allitems[i].LinkToFile
                        });
                    }
                }
            }
        };
        _this._getPeoplePickerItems = function (items) { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                if (items.length > 0) {
                    this.setState({ Owner: items[0].text });
                    this.setState({ OwnerId: items[0].id });
                }
                else {
                    //ID=0;
                    this.setState({ OwnerId: "" });
                    this.setState({ Owner: "" });
                }
                return [2 /*return*/];
            });
        }); };
        _this.onchange = function (e, stateValue) {
            var state = {};
            state[stateValue] = e.target.value;
            _this.setState(state);
        };
        _this.setstatelocal = function (x) {
            var state = {};
            state["LinkToFile"] = x;
            _this.setState(state);
        };
        _this.state = {
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
            LinkToFile: ""
        };
        return _this;
    }
    CrudReact3.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.fetchData()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    CrudReact3.prototype.saveIntoSharePoint = function (file) {
        return __awaiter(this, void 0, void 0, function () {
            var siteUrl, web;
            var _this = this;
            return __generator(this, function (_a) {
                siteUrl = this.props.webURL;
                web = Web(siteUrl);
                if (file.fileAbsoluteUrl == null) {
                    file.downloadFileContent().then(function (r) { return __awaiter(_this, void 0, void 0, function () {
                        var fileUploaded, fileUploaded;
                        return __generator(this, function (_a) {
                            if (r.size <= 10485760) {
                                fileUploaded = web.getFolderByServerRelativeUrl("/Shared%20Documents/").files.add(file.fileName, r, true);
                            }
                            else {
                                fileUploaded = web.getFolderByServerRelativeUrl("/Shared%20Documents/").files.addChunked(file.fileName, r, function (data) { }, true);
                            }
                            return [2 /*return*/];
                        });
                    }); });
                }
                else {
                }
                this.setState({ LinkToFile: siteUrl + "/Shared%20Documents/" + file.fileName });
                return [2 /*return*/];
            });
        });
    };
    CrudReact3.prototype.fetchData = function () {
        return __awaiter(this, void 0, void 0, function () {
            var web, items, html;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        web = Web(this.props.webURL);
                        return [4 /*yield*/, web.lists.getByTitle("Orders").items.select("*", "Owner/Title").expand("Owner/ID").get()];
                    case 1:
                        items = _a.sent();
                        console.log(items);
                        this.setState({ Items: items });
                        return [4 /*yield*/, this.getHTML(items)];
                    case 2:
                        html = _a.sent();
                        this.setState({ HTML: html });
                        return [2 /*return*/];
                }
            });
        });
    };
    CrudReact3.prototype.getHTML = function (items) {
        return __awaiter(this, void 0, void 0, function () {
            var tabledata;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        tabledata = React.createElement("table", { className: styles.table },
                            React.createElement("thead", null,
                                React.createElement("tr", null,
                                    React.createElement("th", null, "Order Number"),
                                    React.createElement("th", null, "Customer Name"),
                                    React.createElement("th", null, "Destination"),
                                    React.createElement("th", null, "Owner"),
                                    React.createElement("th", null, "State"),
                                    React.createElement("th", null, "Link to File"))),
                            React.createElement("tbody", null, items && items.map(function (item, i) {
                                return [
                                    React.createElement("tr", { key: i, onClick: function () { return _this.findData(item.ID); } },
                                        React.createElement("td", null, item.OrderNumber),
                                        React.createElement("td", null, item.CustomerName),
                                        React.createElement("td", null, item.Destination),
                                        React.createElement("td", null, item.Owner.Title),
                                        React.createElement("td", null, item.State),
                                        React.createElement("td", null, FormatDate(item.HireDate)),
                                        React.createElement("td", null,
                                            " ",
                                            React.createElement(Link, { href: item.LinkToFile, target: '_blank' },
                                                "  ",
                                                React.createElement(FontIcon, { iconName: "Dictionary" }),
                                                " ")))
                                ];
                            })));
                        return [4 /*yield*/, tabledata];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    CrudReact3.prototype.SaveData = function () {
        return __awaiter(this, void 0, void 0, function () {
            var web;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        web = Web(this.props.webURL);
                        return [4 /*yield*/, web.lists.getByTitle("Orders").items.add({
                                OwnerId: this.state.OwnerId,
                                HireDate: new Date(this.state.HireDate),
                                Destination: this.state.Destination,
                                OrderNumber: this.state.OrderNumber,
                                CustomerName: this.state.CustomerName,
                                State: this.state.State,
                                LinkToFile: this.state.LinkToFile
                            }).then(function (i) {
                                console.log(i);
                            })];
                    case 1:
                        _a.sent();
                        alert("Created Successfully");
                        this.setState({ Owner: "", HireDate: null, Destination: "", OrderNumber: "", CustomerName: "", State: "", LinkToFile: "" });
                        this.fetchData();
                        return [2 /*return*/];
                }
            });
        });
    };
    CrudReact3.prototype.UpdateData = function () {
        return __awaiter(this, void 0, void 0, function () {
            var web;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        web = Web(this.props.webURL);
                        return [4 /*yield*/, web.lists.getByTitle("Orders").items.getById(this.state.ID).update({
                                OwnerId: this.state.OwnerId,
                                OrderNumber: this.state.OrderNumber,
                                State: this.state.State,
                                CustomerName: this.state.CustomerName,
                                HireDate: new Date(this.state.HireDate),
                                Destination: this.state.Destination,
                            }).then(function (i) {
                                console.log(i);
                            })];
                    case 1:
                        _a.sent();
                        alert("Updated Successfully");
                        this.setState({ Owner: "", HireDate: null, Destination: "", OrderNumber: "", State: "", CustomerName: "", LinkToFile: "" });
                        this.fetchData();
                        return [2 /*return*/];
                }
            });
        });
    };
    CrudReact3.prototype.DeleteData = function () {
        return __awaiter(this, void 0, void 0, function () {
            var web;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        web = Web(this.props.webURL);
                        return [4 /*yield*/, web.lists.getByTitle("Orders").items.getById(this.state.ID).delete()
                                .then(function (i) {
                                console.log(i);
                            })];
                    case 1:
                        _a.sent();
                        alert("Deleted Successfully");
                        this.setState({ Owner: "", HireDate: null, Destination: "", OrderNumber: "", State: "", CustomerName: "", LinkToFile: "" });
                        this.fetchData();
                        return [2 /*return*/];
                }
            });
        });
    };
    CrudReact3.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", null,
            React.createElement("h1", null, "CRUD Operations With ReactJs"),
            this.state.HTML,
            React.createElement("div", { className: styles.btngroup },
                React.createElement("div", null,
                    React.createElement(PrimaryButton, { text: "Create", onClick: function () { return _this.SaveData(); } })),
                React.createElement("div", null,
                    React.createElement(PrimaryButton, { text: "Update", onClick: function () { return _this.UpdateData(); } })),
                React.createElement("div", null,
                    React.createElement(PrimaryButton, { text: "Delete", onClick: function () { return _this.DeleteData(); } }))),
            React.createElement("div", null,
                React.createElement("form", null,
                    React.createElement("div", null,
                        React.createElement(Label, null, "Order Number"),
                        React.createElement(TextField, { defaultValue: ' ', value: this.state.OrderNumber, onChange: function (value) { return _this.onchange(value, "OrderNumber"); } })),
                    React.createElement("div", null,
                        React.createElement(Label, null, "Customer Name"),
                        React.createElement(TextField, { defaultValue: ' ', value: this.state.CustomerName, onChange: function (value) { return _this.onchange(value, "CustomerName"); } })),
                    React.createElement("div", null,
                        React.createElement(Label, null, "Destination"),
                        React.createElement(TextField, { defaultValue: ' ', value: this.state.Destination, multiline: true, onChange: function (value) { return _this.onchange(value, "Destination"); } })),
                    React.createElement("div", null,
                        React.createElement(Label, null, "Owner"),
                        React.createElement(PeoplePicker, { context: this.props.context, personSelectionLimit: 1, 
                            // defaultSelectedUsers={this.state.Owner===""?[]:this.state.Owner}
                            isRequired: false, defaultSelectedUsers: [this.state.Owner ? this.state.Owner : ""], showHiddenInUI: false, principalTypes: [PrincipalType.User], resolveDelay: 1000, ensureUser: true, selectedItems: this._getPeoplePickerItems })),
                    React.createElement("div", null,
                        React.createElement(Label, null, "Date"),
                        React.createElement(DatePicker, { maxDate: new Date(), allowTextInput: false, strings: DatePickerStrings, value: this.state.HireDate, onSelectDate: function (e) { _this.setState({ HireDate: e }); }, ariaLabel: "Select a date", formatDate: FormatDate })),
                    React.createElement("div", null,
                        React.createElement("br", null),
                        React.createElement("br", null),
                        React.createElement(FilePicker, { label: 'Select or upload file', buttonClassName: styles.button, buttonLabel: 'Images', accepts: [".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"], buttonIcon: "FileImage", onSave: this.saveIntoSharePoint.bind(this), onChanged: this.saveIntoSharePoint.bind(this), context: this.props.context }))))));
    };
    return CrudReact3;
}(React.Component));
export default CrudReact3;
export var DatePickerStrings = {
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
export var FormatDate = function (date) {
    console.log(date);
    var date1 = new Date(date);
    var year = date1.getFullYear();
    var month = (1 + date1.getMonth()).toString();
    month = month.length > 1 ? month : '0' + month;
    var day = date1.getDate().toString();
    day = day.length > 1 ? day : '0' + day;
    return month + '/' + day + '/' + year;
};
//# sourceMappingURL=CrudReact3.js.map