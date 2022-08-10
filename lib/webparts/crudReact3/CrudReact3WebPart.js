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
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'CrudReact3WebPartStrings';
import CrudReact3 from './components/CrudReact3';
import "@pnp/sp/lists";
import "@pnp/sp/items";
var CrudReact3WebPart = /** @class */ (function (_super) {
    __extends(CrudReact3WebPart, _super);
    function CrudReact3WebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    CrudReact3WebPart.prototype.render = function () {
        var element = React.createElement(CrudReact3, {
            description: this.properties.description,
            webURL: this.context.pageContext.web.absoluteUrl,
            context: this.context
        });
        ReactDom.render(element, this.domElement);
    };
    CrudReact3WebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(CrudReact3WebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    CrudReact3WebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return CrudReact3WebPart;
}(BaseClientSideWebPart));
export default CrudReact3WebPart;
//# sourceMappingURL=CrudReact3WebPart.js.map