import * as React from "react";
import {MessageBar,MessageBarType} from 'office-ui-fabric-react';

export default class Message extends React.Component {

    constructor(props, context) {
        super(props, context);
    }
    
    handleDismiss = (e) => {
        e.preventDefault();
        const {message} = this.props;
        this.props.onDismiss(message);
    }

    render() {
        const {message} = this.props;
        if(message.type === 'error'){
            return(
                <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                    {message.text}
                </MessageBar>
            );
        }else if(message.type === 'warning') {
            return(
                <MessageBar messageBarType={MessageBarType.warning} isMultiline={false} onDismiss={this.handleDismiss} dismissButtonAriaLabel="Close">
                    {message.text}
                </MessageBar>
            );
        }else if(message.type === 'success') {
            return(
                <MessageBar messageBarType={MessageBarType.success} isMultiline={false} onDismiss={this.handleDismiss} dismissButtonAriaLabel="Close">
                    {message.text}
                </MessageBar>
            );
        }
    }

}