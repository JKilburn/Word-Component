import * as React from "react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";
import Login from "./Login";
import Merge from "./Merge";
import SaveFile from "./SaveFile";
import Messages from "./Messages";
import axios from "axios";

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      fileLocation:'',
      fileName:'',
      objectName:'',
      recordId:'',
      authResult:{},
      contentDocumentLink:{},
      messages:[]
    };
  }

  /** Start App Functions */
  setContentDocumentLink() {
    try{
      let SOQL = 'SELECT Id,ContentDocumentId FROM ContentDocumentLink WHERE LinkedEntityId = \'' + this.state.recordId + '\' AND ContentDocument.Title = \'' + this.state.fileName +'\'';
      let queryURL = decodeURIComponent(this.state.authResult.instance_url) + '/services/data/v46.0/query/?q=' + SOQL;
      let config = {
        headers: {
          "Authorization": 'Bearer ' +  decodeURIComponent(this.state.authResult.access_token),
          "Content-Type": 'application/json'
        }
      } 
      axios.get(queryURL,config).then((response) => {
          var content = response.data.records[0];
          this.setState({
            contentDocumentLink:content
          });
      });
    }catch(ex){
      this.handleMessage(ex.message,'error');
    }
  }

  /* Start Callbacks and Handlers */
  handleAuthentication = (result) => {
    try{
      this.setState({
        authResult:result
      });
      this.setContentDocumentLink();
    }catch(err){
      this.handleMessage(err.message,'error');
    }
    
  }

  handleMessage = (text,type) => {
    let messages = this.state.messages;
    messages.push({text:text,type:type});
    this.setState({
      messages:messages
    });
  }

  handleClearMessage = (errorMessage) => {
    let messages = this.state.messages;
    messages = messages.filter(e => e.text !== errorMessage.text);
    this.setState({
      messages:messages
    });
  }
  /**End Callbacks and Handlers */

  componentDidMount() {
    let fileLocation = Office.context.document.url;
    var filename = fileLocation.replace(/^.*[\\\/]/, '').split('.')[0];
    var filename = filename.substring(0, filename.indexOf('--]')+3);
    var recordDetails = filename.substring(filename.indexOf('[-')+2,filename.indexOf('--]')).split('_');
    this.setState({
      listItems: [
        {icon: "WordDocument",primaryText: "Word documents upload as new version"},
        {icon: "PDF",primaryText: "PDF documents upload as a new document"}
      ],
      fileLocation:fileLocation,
      fileName:filename,
      objectName:recordDetails[0],
      recordId:recordDetails[1]
    });
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/upload-filled.png" message="Please sideload your addin to see app body." />
      );
    }else if(this.state.authResult.access_token != undefined) {
      return (
        <div className="ms-welcome">
          <Header logo="assets/upload-filled.png" title={this.props.title} message="File Upload" />
          <Messages messages={this.state.messages} onDismiss={this.handleClearMessage}></Messages>
          <HeroList message="Upload this file back to Salesforce." items={this.state.listItems}>
            <Merge objectName={this.state.objectName} recordId={this.state.recordId} authResult={this.state.authResult} onMessage={this.handleMessage}>
              <SaveFile 
                authResult={this.state.authResult} 
                onMessage={this.handleMessage} 
                fileName={this.state.fileName} 
                documentId={this.state.contentDocumentLink.ContentDocumentId}
                recordId={this.state.recordId}>
              </SaveFile>
            </Merge>
          </HeroList>
        </div>
      );
    }else{
      return (
        <div className="ms-welcome">
          <Header logo="assets/upload-filled.png" title={this.props.title} message="File Upload" />
          <Messages messages={this.state.messages} onDismiss={this.handleClearMessage}></Messages>
          <HeroList message="Upload this file back to Salesforce" items={this.state.listItems}>
            <Login onSuccess={this.handleAuthentication} onMessage={this.handleMessage}>

            </Login>
          </HeroList>
        </div>
      );
    }

  }

}
