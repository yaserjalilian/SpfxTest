import * as React from 'react';
import { ICrudReact3Props } from './ICrudReact3Props';
import { IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib';
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { ICrudReact3WebPartProps } from '../CrudReact3WebPart';
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
export default class CrudReact3 extends React.Component<ICrudReact3Props, IStates> {
    constructor(props: any);
    componentDidMount(): Promise<void>;
    saveIntoSharePoint(file: IFilePickerResult): Promise<void>;
    fetchData(): Promise<void>;
    findData: (id: any) => void;
    getHTML(items: any): Promise<JSX.Element>;
    _getPeoplePickerItems: (items: any[]) => Promise<void>;
    onchange: (e: any, stateValue: any) => void;
    setstatelocal: (x: any) => void;
    private SaveData;
    private UpdateData;
    private DeleteData;
    render(): React.ReactElement<ICrudReact3WebPartProps>;
}
export declare const DatePickerStrings: IDatePickerStrings;
export declare const FormatDate: (date: any) => string;
//# sourceMappingURL=CrudReact3.d.ts.map