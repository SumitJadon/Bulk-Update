import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';

export default class SharePointServices {

    public static async getAllListDetails(webUrl: string, spHttpClient: SPHttpClient): Promise<any> {
        let requestUrl = webUrl + `/_api/web/lists?$select=Title,Id,EntityTypeName&$filter=BaseTemplate eq 100`;
        const options: ISPHttpClientOptions = {
            headers: { 'odata-version': '3.0' }
        };
        return spHttpClient.get(requestUrl, SPHttpClient.configurations.v1, options)
            .then(async (response: SPHttpClientResponse) => {
                if (response.ok && (response.status === 200)) {
                    const listObject = await response.json();
                    return listObject.value;
                }
                else {
                    console.error(`Error => Code: ${response.status} Message:${response.statusText}`);
                    return null;
                }
            }).catch((error: any) => {
                console.error(`An error occurred while retrieving All List details. Error = ${error}`);
                return null;
            });
    }

    public static async retrieveFileProps(libraryId: string, webUrl: string, spHttpClient: SPHttpClient, folderServerRelativeUrl: string, nextHref?: string): Promise<any> {
        if (!!folderServerRelativeUrl) {
            var data: string = '';
            var queryXML = `<View Scope="RecursiveAll"><ViewFields><FieldRef Name="ID" LookupId="FALSE"/><FieldRef Name="File_x0020_Type" LookupId="FALSE"/></ViewFields><OrderBy><FieldRef Name="ID" Ascending="TRUE"/></OrderBy><Query><Where><Eq><FieldRef Name="FSObjType" /><Value Type="Integer">0</Value></Eq></Where></Query><RowLimit Paged="TRUE">2000</RowLimit></View>`;
            var requestUrl = webUrl + `/_api/web/lists(guid'` + libraryId + `')/RenderListDataAsStream`;
            if (!!nextHref) {
                data = `{'parameters': { '__metadata': {'type': 'SP.RenderListDataParameters'} , 'FolderServerRelativeUrl':'` + folderServerRelativeUrl + `' , 'ViewXml':'` + queryXML + `', 'Paging':'` + nextHref.substring(1) + `'}}`;
            } else {
                data = `{'parameters': { '__metadata': {'type': 'SP.RenderListDataParameters'} , 'FolderServerRelativeUrl':'` + folderServerRelativeUrl + `' , 'ViewXml':'` + queryXML + `'}}`;
            }
            const sphttpClientPostptions: ISPHttpClientOptions = {
                headers: { 'odata-version': '3.0' },
                body: data
            };
            return spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, sphttpClientPostptions)
                .then(async (response: SPHttpClientResponse) => {
                    if (response.ok && (response.status === 200)) {
                        const responseObject = await response.json();
                        return responseObject;
                    }
                    else {
                        console.error(`Error => Code: ${response.status} Message:${response.statusText}`);
                        return null;
                    }
                }).catch((error: any) => {
                    console.error(`An error occurred while retrieving properties of all the Documents from the Document Library. Error = ${error}`);
                    return null;
                });
        }
    }

    public static async retrieveLookUpValues(listId: string, webUrl: string, spHttpClient: SPHttpClient, values: string[]): Promise<any> {
        var queryXML = `<View><ViewFields><FieldRef Name="Title" LookupId="FALSE"/></ViewFields><Query><Where><In><FieldRef Name="Title" /><Values>`;
        values.forEach((value) => {
            queryXML += `<Value Type="Text">` + value + `</Value>`;
        });
        queryXML += `</Values></In></Where></Query><RowLimit>500</RowLimit></View>`;
        var requestUrl = webUrl + `/_api/web/lists(guid'` + listId + `')/RenderListDataAsStream`;
        const sphttpClientPostptions: ISPHttpClientOptions = {
            headers: { 'odata-version': '3.0' },
            body: `{'parameters': { '__metadata': {'type': 'SP.RenderListDataParameters'} , 'ViewXml':'` + queryXML + `'}}`
        };
        return spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, sphttpClientPostptions)
            .then(async (response: SPHttpClientResponse) => {
                if (response.ok && (response.status === 200)) {
                    const responseObject = await response.json();
                    return responseObject;
                }
                else {
                    console.error(`Error => Code: ${response.status} Message:${response.statusText}`);
                    return null;
                }
            })
            .catch((error: any) => {
                console.error(`An error occurred while retrieving KKS Ids from the KKS List. Error = ${error}`);
                return null;
            });
    }

}