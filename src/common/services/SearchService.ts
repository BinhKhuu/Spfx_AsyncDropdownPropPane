import { IDropdown, IDropdownOption }           from "office-ui-fabric-react";
import { sp }                                   from "@pnp/sp";
import                                          "@pnp/sp/search";
import                                          "@pnp/sp/webs";
import                                          "@pnp/sp/lists";
import                                          "@pnp/sp/items";
import                                          "@pnp/sp/content-types";
import { Text } 		                        from '@microsoft/sp-core-library';
import { isEmpty }                              from "@microsoft/sp-lodash-subset";


export class SearchService {
    private siteUrlOptions: IDropdownOption[];
    private webUrlOptions: IDropdownOption[];
    private listUrlOptions: IDropdownOption[];
    public loadSites(): Promise<IDropdownOption[]> {
        console.log("siteurl options", this.siteUrlOptions);
        if(this.siteUrlOptions){
            return Promise.resolve(this.siteUrlOptions);
        }
        else {
            return new Promise<IDropdownOption[]>(async(resolve,reject)=>{
                sp.search({ Querytext: "contentclass:STS_Site", SelectProperties: ["Title", "SPSiteUrl", "Path"], RowLimit: 500, TrimDuplicates: false})
                .then((sites: any) => {
                    let options:IDropdownOption[] = this._CreateURLDropdownOptions(sites.PrimarySearchResults)
                    this.siteUrlOptions = options;
                    resolve(options);
                })
            });
        }
    }

    private _CreateURLDropdownOptions(primarySearchResults): IDropdownOption[] {
        const serverUrl = Text.format("{0}//{1}", window.location.protocol, window.location.hostname);
        let options:IDropdownOption[] = [ {key:"", text: ""}];
        let urlOptions = primarySearchResults.sort().map((result)=>{
            let serverRelativeUrl = !isEmpty(result.Path.replace(serverUrl, '')) ? result.Path.replace(serverUrl, '') : '/';
            return {key: result.Path, text:serverRelativeUrl}
        });
        options = options.concat(urlOptions);
        return options;
    }

    public clearWebsCache(){
        this.webUrlOptions = null;
    }

    public loadWebsFromSite(siteUrl: string): Promise<IDropdownOption[]> {
        if(this.webUrlOptions){
            return Promise.resolve(this.webUrlOptions)
        }
        else {
            return new Promise<IDropdownOption[]>(async(resolve,reject)=>{
                var queryText = Text.format('Path:{0} AND (contentclass:STS_Site OR contentclass:STS_Web)',siteUrl)
                sp.search({Querytext: queryText, SelectProperties: ["Title", "SPSiteUrl", "Path"], RowLimit: 500, TrimDuplicates: false})
                .then((sites: any) => {
                    let options:IDropdownOption[] = this._CreateURLDropdownOptions(sites.PrimarySearchResults);
                    options.some((option)=>{
                        console.log('option',option,siteUrl)
                        if(option.key == siteUrl){
                            var text = option.key.toString().replace(siteUrl,"/");
                            option.text = text;
                        }
                        return option.key == siteUrl;
                    })
                    this.webUrlOptions = options;
                    resolve(options);
                })
            })
        }
    }

    public loadListsFromWeb(webUrl: string): Promise<IDropdownOption[]> {
        console.log("webbbbbbbbbbb",webUrl)



        // if(this.listUrlOptions){
        //     return Promise.resolve(this.listUrlOptions);
        // }
        // else {
            return new Promise<IDropdownOption[]>(async(resolve,reject)=>{
                //var queryText = Text.format('Path:{0}/* AND (contentclass:STS_ListItem_GenericList OR contentclass:STS_List_DocumentLibrary)',webUrl)
                //pnp search will not return list
                //var asdf = sp.web.lists.getByTitle('Tile Items').contentTypes();//7757da70-efd5-47c8-80ef-3a4786e0a48d
                var asdf = sp.web.lists.get();
                var ww = await asdf;
                ww.forEach(()=>{
                    
                })
                //var ww = await asdf.contentTypes();
                console.log("wwwwwwwwwwwwwwwwwwwww",ww)
                var queryText = Text.format('Path:https://blueboxsolutionsdev.sharepoint.com/teams/rc318_8_cdms contentclass:STS_List contentTypeId:0x0100FE14589A34A4B94*',webUrl)//contentTypeId:0x0100FE14589A34A4B94*
                sp.search({Querytext: queryText, RowLimit: 1000, TrimDuplicates: true})
                .then((sites: any) => {
                    console.log("LISTS",sites);
                    let options:IDropdownOption[] = this._CreateURLDropdownOptions(sites.PrimarySearchResults);
                    options.some((option)=>{
                        console.log('option',option,webUrl)
                        if(option.key == webUrl){
                            var text = option.key.toString().replace(webUrl,"/");
                            option.text = text;
                        }
                        return option.key == webUrl;
                    })
                    this.listUrlOptions = options;
                    resolve(options);
                })
            })
        //}
    }

}