import * as React from "react";
import jsforce from "jsforce";
import { Button, ButtonType } from "office-ui-fabric-react";
import {TextField} from 'office-ui-fabric-react/lib/TextField';
import { Dropdown} from 'office-ui-fabric-react/lib/Dropdown';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Spinner, SpinnerType } from "office-ui-fabric-react";

var selectedFileType = 'docx';

export default class SaveFile extends React.Component {

    constructor(props, context) {
        super(props, context);
        this.state = {
            isSaving:false,
            reason:''
        };
    }

    componentDidMount() {
    
    }

    handleError(msg){
        self.setState({
            isSaving:false
        });
        this.props.onMessage(msg,'error');
    }

    save = () => {
        var self = this;
        try{
            self.setState({
                isSaving:true
            });
            this.getFile()
            .then(this.sendFile)
            .then(function(response){
                self.setState({
                    isSaving:false,
                    reason:''
                });
                self.props.onMessage('The file has successfully uploaded with a record id of ' + response,'success');
            })
            .catch(function(e){
                self.handleError(JSON.stringify(e));
                //self.handleError(e);
            });
        }catch(e){
            self.handleError(JSON.stringify(e));
        }
    }

    handleSelectFileType = (event,item) => {
        selectedFileType = item.key;
    }

    handleReasonChange = (event,newValue) => {
        this.setState({ reason: newValue || '' });
    }

    getFile() {
        var self = this;
        return new Promise(function(resolve, reject) {
            Word.run(context => {
                let fileType = 'compressed';
                if(selectedFileType == 'pdf')
                    fileType = 'pdf';
                Office.context.document.getFileAsync(fileType,function (result) {
                    if (result.status == Office.AsyncResultStatus.Succeeded) {
                        var file = result.value;
                        self.getAllSlices(file).then(function(result) {
                            if(result.IsSuccess){
                                resolve(result.Data);
                            }else{
                                reject('There was an issue getting the file slice data.');
                            }
                        });
                    }else {
                        reject('There was an issue getting the file.');
                    }
                });
            });
        });
    }

    getAllSlices(file){
        var isError = false;
        return new Promise(async (resolve, reject) => {
            var documentFileData = [];
            for (var sliceIndex = 0; (sliceIndex < file.sliceCount) && !isError; sliceIndex++) {
                var sliceReadPromise = new Promise((sliceResolve, sliceReject) => {
                    file.getSliceAsync(sliceIndex, (asyncResult) => {
                        if(asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            documentFileData = documentFileData.concat(asyncResult.value.data);
                            sliceResolve({
                                IsSuccess: true,
                                Data: documentFileData
                            });
                        }else{
                            file.closeAsync();
                            sliceReject({
                                IsSuccess: false,
                                ErrorMessage: `Error in reading the slice: ${sliceIndex} of the document`
                            });
                        }
                    });
                });
                await sliceReadPromise.catch((error) => {
                    isError = true;
                });
            } 
            if(isError || !documentFileData.length) {
              reject('Error while reading document. Please try it again.');
              return;
            }
            file.closeAsync();
            resolve({
                IsSuccess: true,
                Data: documentFileData
            });
        });
    }

    sendFile = (content) => {
        var self = this;
        var reason = self.state.reason;
        return new Promise( function( resolve, reject ) {
            try{
                var conn = new jsforce.Connection({
                    instanceUrl : decodeURIComponent(self.props.authResult.instance_url),
                    accessToken : decodeURIComponent(self.props.authResult.access_token)
                });
                let buff = new Buffer(content,'binary');
                let base64data = buff.toString('base64');
                let path = '/services/data/v46.0';
                let docId = self.props.documentId;
                let contentVersionBody = {};
                if(selectedFileType === 'pdf'){
                    contentVersionBody = {
                        'VersionData' : base64data,
                        'PathOnClient': self.props.fileName + '.pdf',
                        'FirstPublishLocationId':self.props.recordId,
                        'ReasonForChange':reason
                    };
                }else{
                    contentVersionBody = {
                        'ContentDocumentId':docId,
                        'VersionData' : base64data,
                        'PathOnClient': self.props.fileName + '.docx',
                        'ReasonForChange':reason
                    };
                }
                
                return conn.requestPost( path + '/composite/', {
                    'allOrNone' : true,
                    'compositeRequest' : [{
                        'method' : 'POST',
                        'url' : path + '/sobjects/ContentVersion',
                        'referenceId' : 'newFile',
                        'body' : contentVersionBody
                    }]
                })
                .then(self.validateCompositeResponse )
                .then( function( response ) {
                    let contentVersionId = null;
                    for ( let i = 0; i < response.compositeResponse.length; i++ ) {
                        if(response.compositeResponse[i].referenceId === 'newFile' ) {
                            contentVersionId = response.compositeResponse[i].body.id;
                            break;
                        }
                    }
                    resolve(contentVersionId);
                }).catch( function( err ) {
                    self.handleError(JSON.stringify(err));
                    reject('something went wrong');
                });
            }catch(err){
                reject('something went wrong');
            }
        });
    }

    validateCompositeResponse = ( response ) => {
        var self = this;
        return new Promise( function( resolve, reject ) {
            try{
                for ( var i = 0; i < response.compositeResponse.length; i++ ) {
                    var body = response.compositeResponse[i].body[0];
                    if ( body && body.errorCode && body.errorCode != 'PROCESSING_HALTED' ) {
                        reject( body.message );
                    }
                }
                resolve(response);
            }catch(ex){
                reject('something went wrong');
            }
        });
    }

    render() {
        let options = [{key:'docx',text: 'Word',data: { icon: 'WordDocument' },isSelected:true},{key:'pdf',text:'PDF',data: { icon: 'PDF' }}];
        if(this.state.isSaving){
            return(
                <div className="ms-Grid" dir="ltr">
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg2">
                            <Spinner type={SpinnerType.large} label="Please wait while the file is uploaded." />
                        </div>
                    </div>
                </div>
            );
        }else{
            return (
                <div className="ms-welcome">
                    <Stack tokens={{ childrenGap: 20 }}>
                        <Dropdown
                            id="fileType"
                            placeholder="Select options"
                            label="Select File Type"
                            options={options}
                            styles={{ dropdown: { width: 300 } }}
                            onChange={this.handleSelectFileType}
                        />
                        <TextField value={this.state.reason} onChange={this.handleReasonChange} id="reason" label="Reason" placeholder="Please provide a reason for updating this file." />
                        <Button className="ms-welcome__action" buttonType={ButtonType.hero} iconProps={{ iconName: "Upload" }} onClick={this.save}>
                            Upload
                        </Button>
                    </Stack>
                </div>
            );
        }
    }

}