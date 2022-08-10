import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import "@pnp/sp/lists";
import "@pnp/sp/items";
export interface ICrudReact3WebPartProps {
    description: string;
}
export default class CrudReact3WebPart extends BaseClientSideWebPart<ICrudReact3WebPartProps> {
    render(): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=CrudReact3WebPart.d.ts.map