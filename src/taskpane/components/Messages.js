import * as React from "react";
import Message from "./Message"

export default class Messages extends React.Component {

    constructor(props, context) {
        super(props, context);
    }

    createMessageElements = () => {
        const {messages} = this.props;
        let messageElements = []
        for (let i = 0; i < messages.length; i++) {
            let message = messages[i];
            messageElements.push(
                <Message message={message} onDismiss={this.props.onDismiss}></Message>
            );
        }
        return messageElements;
    }

    render() {
        const {messages} = this.props;
        if(messages.length > 0) {
            return(
                <div>
                    {this.createMessageElements()}
                </div>
            );
        }else{
            return(<span></span>);
        }
    }

}