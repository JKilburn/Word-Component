import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";

var dialog;
var callbackURL;

export default class Login extends React.Component {
  
    constructor(props, context) {
        super(props, context);
        callbackURL = window.location.origin + '/oauthcallback.html';
    }

    componentDidMount() { 
        //TODO handle persistant behavior
    }

    redirectToAuth = () =>{
        Office.context.ui.displayDialogAsync(
            callbackURL,
            {height: 80, width: 40,displayInIframe: false},
            this.dialogCallback
        );
    }

    dialogCallback = (asyncResult) => {
        if (asyncResult.status == "failed") { 
            var msg = this.convertCodeToText(asyncResult.error.code);
            this.props.onMessage(msg,'warning');
        }else {
            dialog = asyncResult.value;           
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, this.messageHandler); //set a handler for response back from salesforce.
            dialog.addEventHandler(Office.EventType.DialogEventReceived, this.eventHandler); //Set a handler for other event issues with the dialog box.
        }
    }

    messageHandler = (arg) => {
        dialog.close();
        var result = this.expandAuthenticationResults(arg.message);
        if(result.error != null) {
            this.props.onMessage(error.message,'warning');
        }else{
            this.props.onSuccess(result,true);
        }
    }

    expandAuthenticationResults(message) {
        var result = message.split('&').reduce(function (result, item) {
            var parts = item.split('=');
            result[parts[0]] = parts[1];
            return result;
        }, {});
        return result;
    }
    
    eventHandler = (arg) => {
        var errorMessage = this.convertCodeToText(arg.error); //Capture any errors thrown from the dialog.
        this.props.onMessage(errorMessage,'warning');
    }

    convertCodeToText = (errorCode) => {
        var errorMessage;
        switch (errorCode) {
            case 12002:
                errorMessage = 'Cannot load URL, no such page or bad URL syntax.';
                break;
            case 12003:
                errorMessage = 'HTTPS is required.';
                break;
            case 12004:
                errorMessage = 'Domain is not trusted';
                break;
            case 12005:
                errorMessage = 'HTTPS is required';
                break;
            case 12006:
                errorMessage = 'Dialog closed by user';
                break;
            case 12007:
                errorMessage = 'A dialog is already opened.';
                break;
            default:
                errorMessage = 'Something unknown went wrong.';
                break;
        }
        return errorMessage;
    }

    render() {
        const { children } = this.props;
        return (
            <div className="ms-welcome">
                <p className="ms-font-l">
                    To start using this task pane click <b>Login</b>.
                </p>
                <Button className="ms-welcome__action" buttonType={ButtonType.hero} iconProps={{ iconName: "CaretHollow" }} onClick={this.redirectToAuth}>
                    Login
                </Button>
                {children}
            </div>
        );
    }
  
}