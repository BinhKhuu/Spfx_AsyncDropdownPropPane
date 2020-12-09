import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AsyncDropDownPropPaneWebPartStrings';
import AsyncDropDownPropPane from './components/AsyncDropDownPropPane';
import { IAsyncDropDownPropPaneProps } from './components/IAsyncDropDownPropPaneProps';

import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { PropertyPaneAsyncDropdown } from '../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { get, isEmpty, update } from '@microsoft/sp-lodash-subset';


import { SearchService } from '../../common/services/SearchService';
import {AsyncDropDownConstants} from './AsyncDropDownConstants';


export interface IAsyncDropDownPropPaneWebPartProps {
    description: string;
    sites: string;
    webs: string;
    lists: string;
}

export default class AsyncDropDownPropPaneWebPart extends BaseClientSideWebPart<IAsyncDropDownPropPaneWebPartProps> {

     /*
        tutorial here : https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/guidance/build-custom-property-pane-controls
        Webpart Property renderers
    */
   private sitesPropertyDropdown: PropertyPaneAsyncDropdown;
   private websPropertyDropdown: PropertyPaneAsyncDropdown;
   private listPropertyDropdown: PropertyPaneAsyncDropdown;
   //services
   private searchService: SearchService;
   
   protected onInit(): Promise<void> {
       this.searchService = new SearchService();
       return new Promise<void>(async (resolve, reject) => {
           /*
           let serverUrl = Text.format("{0}//{1}", window.location.protocol, window.location.hostname);
           let queryProperties = Text.format("querytext='SPSiteUrl:{0} AND (contentclass:STS_Site OR contentclass:STS_Web)'&selectproperties='Path'&trimduplicates=false&rowLimit=500&Properties='EnableDynamicGroups:true'", serverUrl)
               

           var websites = sp.search({ QueryTemplate: queryProperties, SelectProperties: ["Title", "SPSiteUrl", "WebTemplate"], RowLimit: 500, TrimDuplicates: false
       }).then((res)=>{
               console.log(res)
           })
           
                       var spSites = [];
           sp.search({ Querytext: "contentclass:STS_Site", SelectProperties: ["Title", "SPSiteUrl", "WebTemplate"], RowLimit: 500, TrimDuplicates: false
               }).then((r: any) => {
                   r.PrimarySearchResults.forEach((res)=>{
                       var dropdown = {
                           key:Text.format("{0}", res.SPSiteUrl), 
                           text:Text.format("{0}", res.SPSiteUrl)
                       }
                       // dropdown.key = Text.format("{0}", res.SPSiteUrl);
                       // dropdown.text = Text.format("{0}", res.SPSiteUrl);
                       spSites.push(dropdown);
                   })
                   console.log(spSites)
               })
           
           */

           //console.log("websites", websites)
           
           // var sites = await sp.site.get();
           // console.log("sites", sites)


           // const listContentTypes = await sp.web.lists.getByTitle('Draft')
           //     .contentTypes.get();

           // console.log(listContentTypes)
           // const list = sp.web.lists.getByTitle("Draft");
           // const r = await list.contentTypes();
           // console.log('rrrrr', r);

           // const list1 = sp.web.lists.get().then((data) => {
           //     console.log('data', data)

           // })
           return resolve();

       })
   }

   private _OnSiteChange(propertyPath: string, oldvalue:any,newValue: any): void {
       update(this.properties, propertyPath, (): any => { return newValue; });
       this.sitesPropertyDropdown.properties.selectedKey = this.properties.sites;
       //re render dropdown
       this.sitesPropertyDropdown.render();
       this._ResetDependentPropertyPanes(propertyPath);
       //rerender webpart
       this.render();
   }

   private _OnWebChange(propertyPath: string, oldvalue:any,newValue: any): void {
       update(this.properties, propertyPath, (): any => { return newValue; });
       this.websPropertyDropdown.properties.selectedKey = this.properties.webs;
       this.websPropertyDropdown.properties.disabled = false;
       this.websPropertyDropdown.render();
       this._ResetDependentPropertyPanes(propertyPath);
       this.render();
   }

   private _OnListChange(propertyPath: string, oldvalue:any,newValue: any): void {
       update(this.properties, propertyPath, (): any => { return newValue; });
       this.listPropertyDropdown.properties.selectedKey = this.properties.lists;
       this.websPropertyDropdown.properties.disabled = false;
       this.websPropertyDropdown.render();
       this._ResetDependentPropertyPanes(propertyPath);
       this.render();
   }

   private _ResetDependentPropertyPanes(propertyPath: string){
       if(propertyPath == AsyncDropDownConstants.propertySite){
           this.searchService.clearWebsCache();
           //this.websPropertyDropdown.properties.selectedKey = '';
           update(this.properties,AsyncDropDownConstants.propertyWebs,(): any =>{return ''} )
           this.websPropertyDropdown.properties.selectedKey = this.properties.webs;
           this.websPropertyDropdown.properties.disabled = isEmpty(this.properties.sites);
           this.websPropertyDropdown.render();
       }
   }

   private _OnListTextChange(propertyPath: string){
       alert('changed')
   }


   public render(): void {

       const element: React.ReactElement<IAsyncDropDownPropPaneProps> = React.createElement(

        AsyncDropDownPropPane,
           {
               description: this.properties.description,
               site: this.properties.sites,
               web: this.properties.webs
           }
       );

       ReactDom.render(element, this.domElement);
   }

   protected onDispose(): void {
       ReactDom.unmountComponentAtNode(this.domElement);
   }

   protected get dataVersion(): Version {
       return Version.parse('1.0');
   }


   private loadWebsUrl(): Promise<IDropdownOption[]>{
       return this.searchService.loadWebsFromSite(this.properties.sites)
   }

   private loadSitesUrl(): Promise<IDropdownOption[]>{
       return this.searchService.loadSites();
   }

   private loadListsUrl(): Promise<IDropdownOption[]>{
       return this.searchService.loadListsFromWeb(this.properties.webs);
   }

   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
       //dropdown for sites property
       this.sitesPropertyDropdown = new PropertyPaneAsyncDropdown(AsyncDropDownConstants.propertySite, {
           label: "sites",
           key:AsyncDropDownConstants.propertySite,
           loadOptions:this.loadSitesUrl.bind(this),
           onPropertyChange: this._OnSiteChange.bind(this),
           selectedKey: this.properties.sites
       });
       //dropdown for webs property
       this.websPropertyDropdown = new PropertyPaneAsyncDropdown(AsyncDropDownConstants.propertyWebs,{
           label:"webs",
           key:AsyncDropDownConstants.propertyWebs,
           loadOptions: this.loadWebsUrl.bind(this),
           onPropertyChange: this._OnWebChange.bind(this),
           selectedKey: this.properties.webs ,
           disabled: isEmpty(this.properties.sites)
       });
       //dropdown for lists
       this.listPropertyDropdown = new PropertyPaneAsyncDropdown(AsyncDropDownConstants.propertyLists,{
           label:"lists",
           key:AsyncDropDownConstants.propertyLists,
           loadOptions: this.loadListsUrl.bind(this),
           onPropertyChange: this._OnListChange.bind(this),
           selectedKey: this.properties.lists,
           disabled: isEmpty(this.properties.webs)
       })

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
                               }),
                           ],
                       },
                       {
                           groupName: "Source",
                           groupFields:[
                               this.sitesPropertyDropdown,
                               this.websPropertyDropdown,
                               this.listPropertyDropdown,
                               PropertyPaneTextField('list',{
                                   label: 'listname',
                                   validateOnFocusOut: this._OnListTextChange.bind(this)
                               })
                           ]
                       }
                   ]
               }
           ]
       };
   }
}

