import { WebPartContext } from "@microsoft/sp-webpart-base";
import{SPHttpClient, MSGraphClientV3, SPHttpClientResponse,ISPHttpClientOptions, HttpClient, HttpClientResponse, IHttpClientOptions} from '@microsoft/sp-http'
import{IDropdownOption} from 'office-ui-fabric-react'
import {SPFI, spfi, SPFx as spSPFX} from '@pnp/sp'
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/fields';
import { IListItem } from "../handlelargelist/components/IListItem";
import { stringIsNullOrEmpty } from "@pnp/common";
import {IList, IUser} from '../msgraphapisphttp/components/Msgraphapisphttp';
import { ICustomformwebpartState } from "../customformwebpart/components/ICustomformwebpartState";

export class SPOperations{
    private spContext:SPFI;
    constructor(context:WebPartContext){
        this.spContext=spfi().using(spSPFX(context));
    }
    //#region "SPHttpClient usage functions"
    public getAllLists(context:WebPartContext):Promise<IDropdownOption[]>{
        let listTitles:IDropdownOption[]=[];
        let restApiUrl:string = context.pageContext.web.absoluteUrl+"/_api/web/lists?select=title";
        return new Promise<IDropdownOption[]>(async(resolve,reject)=>{
            context.spHttpClient.get(restApiUrl,SPHttpClient.configurations.v1,{}).then((response:SPHttpClientResponse)=>{
                response.json().then((results:any)=>{
                    console.log(results);
                    results.value.map((result:any)=>{
                        listTitles.push({key:result.Title, text:result.Title});
                    });
                });
                resolve(listTitles);
            },(error:any):void=>{
                reject("error occured "+error);
            });
        });
    }

    public CreateListItems(context:WebPartContext, listTitle:string):Promise<string>{
        let restApiUrl:string=context.pageContext.web.absoluteUrl+"/_api/web/lists/getByTitle('"+listTitle+"')/items";
        const body:string = JSON.stringify({Title:"New item created"});
        const options:ISPHttpClientOptions = {
            headers:{
            "accept":"application/json;odata=nometadata", 
            "content-type":"application/json;odata=nometadata",
            "odata-version":""}, 
            body:body};
        return new Promise<string>(async (resolve,reject)=>{
            context.spHttpClient.post(restApiUrl,SPHttpClient.configurations.v1,options).then((response:SPHttpClientResponse)=>{
                if (response.ok) {
                    response.json().then((responseJSON) => {
                      console.log(responseJSON);
                      resolve(`Item created successfully with ID: ${responseJSON.ID}`);
                    });
                  } else {
                    response.json().then((responseJSON) => {
                      console.log(responseJSON);
                      resolve(`Something went wrong! Check the error in the browser console.`);
                    });
                  }
            }).catch(error => {
                console.log(error);
            });
        }); 
    }

    public UpdateListItems(context:WebPartContext, listTitle:string):Promise<string>{
        let restApiUrl:string=context.pageContext.web.absoluteUrl+"/_api/web/lists/getByTitle('"+listTitle+"')/items";
        const body:string=JSON.stringify({
            Title:"Updated Item"
        });
        return new Promise<string>(async(resolve,reject)=>{
            this.getLatestItemId(context,listTitle).then((itemId:number)=>{
                context.spHttpClient.post(restApiUrl+"("+itemId+")", SPHttpClient.configurations.v1,{headers:
                    {"accept":"application/json;odata=nometadata", 
                    "content-type":"application/json;odata=nometadata",
                    "odata-version":"",
                    "IF-MATCH":"*",
                    "X-HTTP-METHOD":"MERGE"
                    },
                    body:body,
                }).then((response:SPHttpClientResponse)=>{
                    resolve("Item with id " + itemId + " updated successfully")
                },(error:any)=>{
                    reject("Error occured " + error);
                })
            })
        })
    }

    public DeleteListItems(context:WebPartContext, listTitle:string):Promise<string>{
        let restApiUrl:string=context.pageContext.web.absoluteUrl+"/_api/web/lists/getByTitle('"+listTitle+"')/items";
        return new Promise<string>(async(resolve,reject)=>{
            this.getLatestItemId(context,listTitle).then((itemId:number)=>{
                context.spHttpClient.post(restApiUrl+"("+itemId+")", SPHttpClient.configurations.v1,{headers:
                    {"accept":"application/json;odata=nometadata", 
                    "content-type":"application/json;odata=nometadata",
                    "odata-version":"",
                    "IF-MATCH":"*",
                    "X-HTTP-METHOD":"DELETE"
                    }
                }).then((response:SPHttpClientResponse)=>{
                    resolve("Item with id " + itemId + " deleted successfully")
                },(error:any)=>{
                    reject("Error occured " + error);
                })
            })
        })
    }

    public getLatestItemId(context:WebPartContext, listTitle:string):Promise<number>{
        let restApiUrl:string=context.pageContext.web.absoluteUrl+"/_api/web/lists/getByTitle('"+listTitle+"')/items/?$orderby=ID desc&$top=1&select=id";
        return new Promise<number>(async(resolve,reject)=>{
            context.spHttpClient.get(restApiUrl,SPHttpClient.configurations.v1,{headers:{"accept":"application/json;odata=nometadata", 
            "content-type":"application/json;odata=nometadata",
            "odata-version":""}}).then((response:SPHttpClientResponse)=>{
                if (response.ok) {
                    response.json().then((responseJSON) => {
                      resolve(responseJSON.value[0].ID);
                    },
                    (error:any)=>{
                        reject("Error occured " + error);
                    });
                  } else {
                    response.json().then((responseJSON) => {
                      resolve(-1);
                    });
                  }
            })
        });
    }

    public CallPowerAutomate(context:WebPartContext, listTitle:string, flowUrl:string):Promise<HttpClientResponse>{
        const body:string = JSON.stringify({Title:"New item created from flow"});
        const options:IHttpClientOptions = {
            headers:{
            "content-type":"application/json;odata=nometadata",
            "accept":"application/json;odata=nometadata",
            "odata-version":""}, 
            body:body};
        return new Promise<HttpClientResponse>(async (resolve,reject)=>{
            context.httpClient.post(flowUrl,HttpClient.configurations.v1,options).then((response:HttpClientResponse)=>{
                console.log(response);
                response.json().then((responseJSON:any)=>{
                    resolve(responseJSON.value);
                },(error:any)=>{
                    reject("Error occured " + error);
                });
            }).catch(error => {
                console.log(error);
            });
        }); 
    }
    //#endregion

    //#region "PnP functions"
    public getAllListsByPnP(context:WebPartContext):Promise<IDropdownOption[]>{
        let listTitles: IDropdownOption[]=[];
        const sp=spfi().using(spSPFX(context));
        return new Promise<IDropdownOption[]>(async(resolve,reject)=>{
            sp.web.lists.select("Title")().then((results:any)=>{
                results.map((result:any)=>{
                    listTitles.push({key:result.Title, text:result.Title});
                });
                resolve(listTitles);
            },(error:any)=>{
                reject("error occured "+error);
            })
        });
    }

    public CreateListItemsByPnP(context:WebPartContext, listTitle:string):Promise<string>{
        const sp=spfi().using(spSPFX(context));
        return new Promise<string>(async(resolve,reject)=>{
            sp.web.lists.getByTitle(listTitle).items.add({Title:"Pnp Item"}).then((results:any)=>{
                resolve("Item with id "+ results.data.ID + "added successfully");
            },(error:any)=>{
                reject("error occured "+error);
            })
        });
    }

    public DeleteListItemByPnP(context:WebPartContext, listTitle:string):Promise<string>{
        const sp=spfi().using(spSPFX(context));
        return new Promise<string>(async(resolve,reject)=>{
            this.getLatestItemIdByPnP(context,listTitle).then((itemid:number)=>{
                sp.web.lists.getByTitle(listTitle).items.getById(itemid).delete().then((results:any)=>{
                    resolve("Item with id "+ itemid + " deleted succesfully.");
                })
            })
        });
    }

    public UpdateListItemByPnP(context:WebPartContext, listTitle:string):Promise<string>{
        const sp=spfi().using(spSPFX(context));
        return new Promise<string>(async(resolve,reject)=>{
            this.getLatestItemIdByPnP(context,listTitle).then((itemid:number)=>{
                sp.web.lists.getByTitle(listTitle).items.getById(itemid).update({Title:"Updated PnP item"}).then((results:any)=>{
                    resolve("Item with id "+ itemid + " updated succesfully.");
                })
            })
        });
    }

    public getLatestItemIdByPnP(context:WebPartContext, listTitle:string):Promise<number>{
        const sp=spfi().using(spSPFX(context));
        return new Promise<number>(async(resolve,reject)=>{
            sp.web.lists.getByTitle(listTitle).items.select("ID").orderBy("ID", false).top(1)().then((response:any)=>{
                resolve(response[0].Id);
            })
        })
    }

    public SetPeoplePicker(context:WebPartContext, listTitle:string, users:any[]):Promise<string>{
        const sp=spfi().using(spSPFX(context));
        return new Promise<string>(async(resolve,reject)=>{
            sp.web.lists.getByTitle(listTitle).items.add({Title:"PeoplePickerEntry", PnPPeoplePickerId:users}).then((results:any)=>{
                resolve("Item with id "+ results.data.ID + "added successfully");
            },(error:any)=>{
                reject("error occured "+error);
            })
        });
    }

    public SetTaxonomyControlValue(context:WebPartContext, listTitle:string, values:any, multiValues:any):Promise<string>{
        const sp=spfi().using(spSPFX(context));
        let isMultiValues=false;
        if(values && values !== null && values !== undefined && values.length >0){
            isMultiValues=false;
        }
        else if(multiValues && multiValues !== null && multiValues !== undefined){
            isMultiValues=true;
        }
        return new Promise<string>(async(resolve,reject)=>{
            if(!isMultiValues){
                sp.web.lists.getByTitle(listTitle).items.add({Title:"SingleValue Taxonomy Picker", 
                        singlevaluetax:{ 
                                        Label:values[0].name,
                                        TermGuid:values[0].key,
                                        WssId:-1
                                    }
                                    }).then((results:any)=>{
                    resolve("Item with id "+ results.data.ID + "added successfully");
                },(error:any)=>{
                    reject("error occured "+error);
                })
            }
            else{
                let multTaxValues:string=""; 
                multiValues.map((mVal:any)=>{
                    multTaxValues +=`-1;#${mVal.name}|${mVal.key};#`
                })
                sp.web.lists.getByTitle(listTitle).fields.getByTitle('multvaluetax_0')().then((field)=>{
                    const fldInternalName:string= field.InternalName;
                    const data:any={Title:"Multi value taxonomy field"};
                    data[fldInternalName]=multTaxValues;
                    sp.web.lists.getByTitle(listTitle).items.add(data).then((results:any)=>{
                        resolve("Item with id "+ results.data.ID + "added successfully");
                        },(error:any)=>{
                            reject("error occured "+error);
                        })
                })
            }
        });
    }
    
    //#endregion
    
    //#region "Handle Large Lists though PnP"
    public async getMoreThan5KListItemsWithoutWhereClause(listTitle:string):Promise<IListItem[]>{
        const results:IListItem[]=[];
        return new Promise<IListItem[]>(async(resolve,reject)=>{
            this.spContext.web.lists.getByTitle(listTitle).items.getAll().then((items:any)=>{
                items.map((item:any, index:any)=>{
                    results.push({title:item.Title});
                    console.log(index);
                });
                console.log("Result count: "+results.length);
                resolve(results);
            },(error:any)=>{
                reject("error occured "+error);
            });
        });
    }

    public async getMoreThan5KListItemsWithWhereClause(listTitle:string):Promise<IListItem[]>{
        var results:IListItem[]=[];
        const pageSize:number=5000;
        return new Promise<IListItem[]>(async(resolve,reject)=>{
            const asyncFunctions:any[] = [];
            this.getMaximumId(listTitle).then(async (maxId)=>{
                for (let i=0;i<Math.ceil(maxId/pageSize);i++){
                    let minId=i*pageSize+1;
                    let maxId=(i+1)*pageSize;
                    let resolvePagedListItems = () =>{
                        return new Promise(async (resolve)=>{
                            let pagedItems:IListItem[] = await this.getItemsforEachIteration(listTitle, minId, maxId);
                            resolve(pagedItems);
                        })
                    }
                    asyncFunctions.push(resolvePagedListItems());
                }
                const allResults = await Promise.all(asyncFunctions);
                for(let k=0;k<allResults.length;k++){
                    allResults[k].map((item:any, index:any)=>{
                        results.push({title:item.title});
                    });
                }
                resolve(results);
            });
        });
    }

    private async getItemsforEachIteration(listTitle:string, minId:number, maxId:number):Promise<IListItem[]>{
        const pageResults:IListItem[]=[];
        const camlQuery=`<View>
                            <Query>
                                <Where>
                                    <And>
                                        <And>
                                            <Geq><FieldRef Name='ID'></FieldRef><Value Type='Number'>` +
                                                minId+
                                            `</Value></Geq>
                                            <Leq><FieldRef Name='ID'></FieldRef><Value Type='Number'>` +
                                                maxId+
                                            `</Value></Leq>
                                        </And>
                                        <Eq>
                                            <FieldRef Name='field_5' />
                                            <Value Type='Text'>New York</Value>
                                        </Eq>
                                    </And>
                                </Where>
                            </Query>
                        </View>`;
        return new Promise<IListItem[]>((resolve,reject)=>{
            this.spContext.web.lists.getByTitle(listTitle).getItemsByCAMLQuery({ViewXml:camlQuery}).then((items:any)=>{
                console.log("Paged Result : " + items.length);
                items.map((item:any, index:any)=>{
                    pageResults.push({title:item.Title});
                });
                resolve(pageResults);
            });
        })
    }

    private async getMaximumId(listTitle:string):Promise<number>{
        let maxId:number;
        return new Promise<number>((resolve,reject)=>{
            this.spContext.web.lists.getByTitle(listTitle).items.orderBy("Id",false).top(1).select("Id")().then((results)=>{
                if(results.length>0){
                    maxId=results[0].Id;
                }
                resolve(maxId);
            },(error:any)=>{
                reject("error occured "+error);
            })
        })
    }
    //#endregion

    //#region "Handle Large Lists though Rest API"
    public async getMoreThan5KListItemsRestAPIWithoutWhereClause(context:WebPartContext, listTitle:string, url?:string):Promise<IListItem[]>{
        const results:IListItem[]=[];
        return new Promise<IListItem[]>(async(resolve,reject)=>{
            let iterativeItemsFunc = (context:WebPartContext, listTitle:string, url?:string) =>{
                let restAPIUrl=context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('" + listTitle + "')/items?$orderby=Id%20desc";
                if(!stringIsNullOrEmpty(url)){
                    restAPIUrl=url;
                }
                context.spHttpClient.get(restAPIUrl, SPHttpClient.configurations.v1).then((response:SPHttpClientResponse)=>{
                    response.json().then((responseJson)=>{
                        responseJson.value.map((item:any, index:any)=>{
                        results.push({title:item.Title});
                        });
                        console.log("Result count: "+results.length);
                        if(responseJson['@odata.nextLink']){
                            iterativeItemsFunc(context, listTitle,responseJson['@odata.nextLink'] );
                        }
                        else{
                            resolve(results);
                        }
                    });
                });
            }
            iterativeItemsFunc(context, listTitle);
        });
    }

    public async getMoreThan5KListItemsRestAPIWithWhereClause(context:WebPartContext, listTitle:string):Promise<IListItem[]>{
        var results:IListItem[]=[];
        const pageSize:number=5000;
        return new Promise<IListItem[]>(async(resolve,reject)=>{
            // Array to hold async calls  
            const asyncFunctions:any[] = [];
            this.getRestAPIMaximumId(context,listTitle).then(async (maxId)=>{
                for (let i=0;i<Math.ceil(maxId/pageSize);i++){
                    let minId=i*pageSize+1;
                    let maxId=(i+1)*pageSize;
                    let resolvePagedListItems = () => {  
                        return new Promise(async (resolve) => { 
                            let pagedItems:IListItem[]= await this.getItemsRestAPIforEachIteration(context, i, pageSize, listTitle, minId, maxId);
                            resolve(pagedItems);
                        });
                    };
                    asyncFunctions.push(resolvePagedListItems());
                }
                const allResults = await Promise.all(asyncFunctions);
                for(let k=0;k<allResults.length;k++){
                    allResults[k].map((item:any, index:any)=>{
                        results.push({title:item.Title});
                    });
                }
                resolve(results);
            });
        });
    }

    private async getItemsRestAPIforEachIteration(context:WebPartContext, index:number, pageSize:number, listTitle:string, minId:number, maxId:number):Promise<IListItem[]>{
        let restAPIUrl=context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('" + listTitle + "')/items?$skiptoken=Paged=TRUE%26p_ID="+
            (index * pageSize +1)+"&$top="+pageSize+"&$select=Title, field_5";
        return new Promise<IListItem[]>((resolve,reject)=>{
            context.spHttpClient.get(restAPIUrl, SPHttpClient.configurations.v1).then((response:SPHttpClientResponse)=>{
                response.json().then((responseJson)=>{
                    console.log("Paged Result : " + responseJson.value.length);
                    resolve(responseJson.value);
                })
            });
        })
    }

    private getRestAPIMaximumId(context: WebPartContext,listTitle:string):Promise<number>{
        let maxId:number;
        let restAPIUrl=context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('" + listTitle + "')/items?$orderby=Id%20desc&$top=1&$select=ID";
        return new Promise<number>((resolve,reject)=>{
            context.spHttpClient.get(restAPIUrl, SPHttpClient.configurations.v1).then((response:SPHttpClientResponse)=>{
                response.json().then((responseJson)=>{
                    maxId=responseJson.value[0].Id;
                    resolve(maxId);
                })
            },(error:any)=>{
                reject("error occured "+error);
            })
        })
    }
    //#endregion

    //#region "Custom form webpart functions"
    public SubmitFormData(listTitle:string,fieldTitle:string, multiCourses:string, values:ICustomformwebpartState):Promise<string>{
        return new Promise<string>(async(resolve,reject)=>{
            this.spContext.web.lists.getByTitle(listTitle).fields.getByTitle(fieldTitle)().then((field)=>{
                const fldInternalName:string= field.InternalName;
                const data:any={
                    Title:"Custom list form data",
                    AvailableOnWeekdays: values.availability,
                    EmployeesId: values.employees,
                    Mobile: values.mobile,
                    Address: values.address,
                    Email: values.email,
                    ManagerApproval: values.mgrApproval,
                    Course:{
                        Label:values.courses[0].name,
                        TermGuid:values.courses[0].key,
                        WssId: -1
                    }
                };
                data[fldInternalName]=multiCourses;
                this.spContext.web.lists.getByTitle(listTitle).items.add(data).then((results)=>{
                    resolve("Item with id "+ results.data.ID + "added successfully");
                },(error:any)=>{
                    reject("error occured "+error);
                });
            })
        });
    }
    //#endregion

    //#region "MSGraphClient using SPHttp"
    public getUsers(context:WebPartContext):Promise<IUser[]>{
        let users:IUser[]=[];
        return new Promise<IUser[]>((resolve,reject)=>{
            context.msGraphClientFactory.getClient("3").then((msGraphClient: MSGraphClientV3) => {
                msGraphClient.api("users").version("v1.0").select("displayName,mail").get((err, res) => {
                    if (err) {
                        console.log("Error occured " + err);
                    }
                    else {
                        res.value.map((result: any) => {
                        users.push({ displayName: result.displayName, mail: result.mail });
                        });
                    }
                    resolve(users);
                });
            });
        });
     }

     /*public async getAllItemsFromList(context:WebPartContext, listName:string):Promise<IList[]>{
        let listItems:IList[]=[];
        return new Promise<IList[]>(async (resolve,reject)=>{
            context.msGraphClientFactory.getClient("3").then((msGraphClient: MSGraphClientV3) => {
                msGraphClient
                .api("sites/ngsp.sharepoint.com,f8f8ceda-ee50-46f0-a164-10369d2ade07,2a312457-41c7-4dca-b577-ab776822e8ad/lists/fd119412-b2f4-433d-bb38-f2ff682ec23e/items")
                .expand("fields($select=Title,field_2)")
                .version("v1.0").get(async (err, res) => {
                    if (err) {
                        console.log("Error occured " + err);
                    }
                    else {
                        res.value.map((result: any) => {
                            listItems.push({ title: result.fields.Title, mail: result.fields.field_2 });
                        });
                    }
                    if(res["@odata.nextLink"] !=null || res["@odata.nextLink"] !=undefined){
                        let recursiveCall=(res:any)=>{
                            let nextLink = new URL(res["@odata.nextLink"]);
                            let skipToken = nextLink.searchParams.get("$skipToken");
                            if(skipToken == null){
                                skipToken = nextLink.searchParams.get("$skiptoken");
                            }
                            return new Promise<IList[]>(async (resolve)=>{
                                context.msGraphClientFactory.getClient("3").then((msGraphClient: MSGraphClientV3) => {
                                    msGraphClient
                                    .api("sites/ngsp.sharepoint.com,f8f8ceda-ee50-46f0-a164-10369d2ade07,2a312457-41c7-4dca-b577-ab776822e8ad/lists/fd119412-b2f4-433d-bb38-f2ff682ec23e/items")
                                    .expand("fields($select=Title,field_2)")
                                    .version("v1.0").skipToken(skipToken).get(async (err, res) => {
                                        if (err) {
                                            console.log("Error occured " + err);
                                        }
                                        else {
                                            res.value.map((result: any) => {
                                                listItems.push({ title: result.fields.Title, mail: result.fields.field_2 });
                                            });
                                        }
                                        if(res["@odata.nextLink"] !=null || res["@odata.nextLink"] !=undefined){
                                            resolve(listItems);
                                            await recursiveCall(res);
                                        }
                                        else{
                                            //resolve(listItems);
                                            return;
                                        }
                                    });
                                })
                            })
                        };
                        await recursiveCall(res);
                    }
                    else{
                        resolve(listItems);
                    }
                    resolve(listItems);
                });
            });
        });
     }*/

     public async getAllItemsFromList(context:WebPartContext, listName:string):Promise<IList[]>{
        let listItems:IList[]=[];
        let recursiveCall=(nextlink?:string)=>{
            let skipToken:string=null;
            if(nextlink !==null && nextlink !==undefined && nextlink !==""){
                let nextLink = new URL(nextlink);
                skipToken = nextLink.searchParams.get("$skipToken");
                if(skipToken == null){
                    skipToken = nextLink.searchParams.get("$skiptoken");
                }
            }
            return new Promise<IList[]>(async (resolve)=>{
                context.msGraphClientFactory.getClient("3").then((msGraphClient: MSGraphClientV3) => {
                    msGraphClient
                    .api("sites/ngsp.sharepoint.com,f8f8ceda-ee50-46f0-a164-10369d2ade07,2a312457-41c7-4dca-b577-ab776822e8ad/lists/fd119412-b2f4-433d-bb38-f2ff682ec23e/items")
                    .expand("fields($select=Title,field_2)")
                    .version("v1.0").skipToken(skipToken).get(async (err, res) => {
                        if (err) {
                            console.log("Error occured " + err);
                        }
                        else {
                            res.value.map((result: any) => {
                                listItems.push({ title: result.fields.Title, mail: result.fields.field_2 });
                            });
                        }
                        if(res["@odata.nextLink"] !==null &&  res["@odata.nextLink"] !==undefined){
                            resolve(await recursiveCall(res["@odata.nextLink"]));
                        }
                        else{
                            resolve(listItems);
                        }
                    });
                })
            })
        };
        return new Promise<IList[]>(async (resolve)=>{
            await recursiveCall();
            resolve(listItems);
        });
     }
    //#endregion
}