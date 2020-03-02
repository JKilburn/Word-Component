import * as React from "react";
import axios from "axios";
import { Spinner, SpinnerType } from "office-ui-fabric-react";

export default class Merge extends React.Component {

    constructor(props, context) {
        super(props, context);
        this.state = {
            merged:false
        };
    }

    handleError(msg){
        this.setState({merged:true});
        this.props.onMessage(msg,'error');
    }

    componentDidMount() {
        this.mergeFile();
    }
    
    mergeFile() {
        var self = this;
        Word.run(async context => {
            var searchResults = context.document.body.search('[{][{][{]*[}][}][}]', {matchWildCards: true});
            context.load(searchResults, 'text');
            return context.sync().then(function (){
                self.getMergeFields(searchResults)
                .then(self.getMergeContent)
                .then(self.replaceText)
                .then(function(){
                    self.setState({merged:true});
                })
                .catch(function(e){
                    self.handleError(e);
                });
            });
        });
    }

    getMergeFields(searchResults) {
        var self = this;
        return new Promise(function(resolve, reject) {
            try{
                var mergeFields = ['Id'];
                for (var i = 0; i < searchResults.items.length; i++) {
                    mergeFields.push(searchResults.items[i].text.replace(/{/g,"").replace(/}/g,""));
                }
                var uniqueFieldList = Array.from(new Set(mergeFields));
                var fieldQuery = uniqueFieldList.join(", ");
                resolve(fieldQuery);
            }catch(e) {
                reject(e);
            }
        });
    }

    getMergeContent = ( response ) => {
        const {objectName,recordId,authResult } = this.props;
        var self = this;
        return new Promise(function(resolve, reject) {
            try{
                if(objectName == null || recordId == null)
                    throw 'Unable to detect object name or record id in document title.';
                let SOQL = 'SELECT ' + response + ' FROM ' + objectName + ' WHERE Id = \'' + recordId + '\'';
                let queryURL = decodeURIComponent(authResult.instance_url) + '/services/data/v46.0/query/?q=' + SOQL;
                let config = {
                    headers: {
                    "Authorization": 'Bearer ' +  decodeURIComponent(authResult.access_token),
                    "Content-Type": 'application/json'
                    }
                }  
                axios.get(queryURL,config).then((response) => {
                    var content = response.data.records[0];
                    resolve(content);
                },(e) => {
                    throw 'There was an issue getting merge content. ';
                }); 
            }catch(e){
                reject(e);
            }
        });
    }

    replaceText = (mergeContent) =>{
        var self = this;
        return new Promise(function(resolve,reject){
            Word.run(async context => {
                var searchResults = context.document.body.search('[{][{][{]*[}][}][}]', {matchWildCards: true});
                context.load(searchResults, 'text');
                return context.sync().then(function () {
                    for (var i = 0; i < searchResults.items.length; i++) {
                        var key = searchResults.items[i].text.replace(/{/g,"").replace(/}/g,"");
                        var value = self.getValue(mergeContent,key);
                        searchResults.items[i].insertText(value, "Replace");
                    }
                    resolve(true);
                    return context.sync();
                });
            }).catch(function (error) {
                reject(error);
            });
        });
    }

    getValue(data,key) {
        var parts = key.split('.'); //Seperate key by dot notation
        var value = data;
        for(var i = 0; i < parts.length; i++){
            var part = parts[i];
            value = value[part];
        }
        return value;
    }

    render() {
        const { children,title } = this.props;
        if(this.state.merged){
            return(
                <div>{children}</div>
            );
        }else{
            return(
                <div className="ms-Grid" dir="ltr">
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg2">
                            <Spinner type={SpinnerType.large} label="Updating document merge fields." />
                        </div>
                    </div>
                </div>
            );
        }
    }

}