import * as React from 'react'
import {
    IBaseFloatingPicker,
    IBaseFloatingPickerSuggestionProps,
    FloatingPeoplePicker,
    SuggestionsStore,
} from 'office-ui-fabric-react/lib/FloatingPicker';
import { ICalloutProps} from 'office-ui-fabric-react/lib/Callout';
import { IPersonaProps, IPersona } from 'office-ui-fabric-react/lib/Persona';
import { IInputs } from '../generated/ManifestTypes';
import {  mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { insertSpanAtCursorPosition, insertTextAtCursorPosition, saveCursorPostion, cursorPositionRestore } from './utils';

const classNames = mergeStyleSets({
    item: {
        borderRadius: '5px',
        border: 'solid',
        borderColor:'#eee',
        color: 'rgb(0, 0, 0)',
        padding: '5px',
        maxWidth: '100%',
        outline: 0,
        fontFamily: 'SegoeUI',
        selectors: {
            '&:focus': {
                border: '1px solid black',
                color: 'black',
                backgroundColor: '#eee',
            }
        }
    }
});


export interface IMentionProps {
    context: ComponentFramework.Context<IInputs>,
    people?: IPersonaProps[],
    value?: string,
    allowedNumberOfCharacters: number,
    notifyOutputChanged: () => void,
    formatNumber: (n: number) => string,
    sendEmail?: string,
    fromUser?: string,
    subject?: string,
    description?: string
}

export interface IMentionControlState extends React.ComponentState {
}
export default class MentionEditorControl extends React.Component<IMentionProps, IMentionControlState>{
    private _callOutProps: ICalloutProps;
    private _picker = React.createRef<IBaseFloatingPicker>();
    private _suggestionProps: IBaseFloatingPickerSuggestionProps;
    private _updatedValue: any;
    private _contendEditableRef: React.RefObject<HTMLDivElement>;

    constructor(props: IMentionProps) {
        super(props);
        this._contendEditableRef = React.createRef<HTMLDivElement>();
        this._updatedValue = props.value;
        this.state = {
            peopleList: props.people,
            value: props.value,
            selectedUsers: [],
        };

        this._callOutProps = {
            target: '#calloutSpan',
        }
        this._suggestionProps = {

            footerItemsProps: [
                {
                    renderItem: () => {
                        return <>Showing {this._picker.current ? this._picker.current.suggestions.length : 0} results</>;
                    },
                    shouldShow: () => {
                        return !!this._picker.current && this._picker.current.suggestions.length > 0;
                    },
                },
            ],
        };
    }

    render() {
        return (
            <>
                <div>
                    <div className={classNames.item}
                        contentEditable
                        ref={this._contendEditableRef}
                        id="contentEditableRef"
                        onKeyUp={this._handleEditorKeyUp}
                        onInput={this._onChange}
                        suppressContentEditableWarning={true}
                    >{this._updatedValue}</div>
                    {this._renderFloatingPicker()}
                    <div style={{ textAlign: "right" }}>
                        <hr />
                        <label>
                            {this.getFormattedCharactersRemaining()}
                        </label>
                    </div>
                </div>
            </>
        )
    }
   
    private _onChange = (e: React.FormEvent) => {
        var savedSelection = saveCursorPostion(document.getElementById('contentEditableRef'));
        this.props.notifyOutputChanged();
        this.setState({
            value: e.currentTarget.textContent || ""
        }, () => {
            //restore caret position(s)
            cursorPositionRestore(document.getElementById('contentEditableRef'), savedSelection);
        })
    }

    private getRemainingNumberOfCharacters = (): number => {
        return this.props.allowedNumberOfCharacters - (this.getValue()?.length || 0);
    }
    private getFormattedCharactersRemaining = (): string => this.props.formatNumber(this.getRemainingNumberOfCharacters());
    private _renderFloatingPicker(): JSX.Element {


        return (
            <>
                <FloatingPeoplePicker pickerCalloutProps={this._callOutProps}
                    suggestionsStore={new SuggestionsStore<IPersonaProps>()}
                    className="floatingPickerDiv"
                    getTextFromItem={(persona: IPersonaProps) => persona.text || ''}
                    pickerSuggestionsProps={this._suggestionProps}
                    key="normal"
                    componentRef={this._picker}
                    onChange={this._onPickerChange}
                    suggestionItems={this.state.peopleList}
                /></>);
    }


    private _onPickerChange = (selectedSuggestion: IPersonaProps): void => {
        insertTextAtCursorPosition(selectedSuggestion.primaryText || '');
        if (this._picker.current) {
            this._picker.current.hidePicker();
        }
        if (this.props.sendEmail == "0") {
            this.sendEmail(selectedSuggestion.optionalText);
        }
    };

    
    getValue = (): string | undefined => this.state.value || "";
    setValue = (value: string | null) => {
        // Mediate whether the value should be updated in state.
        // Without this check, the value could be overwritten by an old value received by the PCF framework.
        if (this.state.active) return;
        this.setValueInternal(value);
    }
    private setValueInternal = (value: string | null): void => {
        this.setState({ value }, this.props.notifyOutputChanged);
    }

    private sendEmail(userId?: string) {
        // @ts-ignore;
        var entityName = Xrm.Page.data.entity.getEntityName(); var entityId = Xrm.Page.data.entity.getId().replace('{', '').replace('}', '');      
        // @ts-ignore;
        var recordUrl = `${Xrm.Page.context.getClientUrl()}/main.aspx?forceUCI=1&newWindow=true&pagetype=entityrecord&etn=${entityName}&id=${entityId}`;
        // @ts-ignore;
        let fromUser = this.props.fromUser != "" ? this.props.fromUser : Xrm.Page.context.getUserId().replace('{', '').replace('}', '');
        var url = "<a href='"+recordUrl+"'> Click Here</a>";
        let description = `${this.props.description} ${url}`;
        var data =
        {
            "subject": this.props.subject,
            "description": description,

            "email_activity_parties": [
                {
                    "partyid_systemuser@odata.bind": "/systemusers(" + fromUser + ")",
                    "participationtypemask": 1  ///From Email
                },
                {
                    "partyid_account@odata.bind": "/systemusers(" + userId + ")",
                    "participationtypemask": 2  ///To Email
                }]
        }

        //@ts-ignore
        Xrm.WebApi.createRecord("email", data).then(
            function success(result: { id: string; }) {
                console.log("Email created with ID: " + result.id);
                sendNotification(result.id);
            },
            function (error: { message: any; }) {
                console.log(error.message);
            }
        );
    }


    private _handleEditorKeyUp = (e: React.KeyboardEvent) => {
        var windowSelection = window.getSelection();
        if (windowSelection != null) {
            var anchorNode = windowSelection.anchorNode?.nodeValue;
            var focusOffset = windowSelection.focusOffset;
            const lastCharacter = anchorNode && anchorNode[focusOffset - 1]

            if (lastCharacter === '@') {
                insertSpanAtCursorPosition("calloutSpan");
                if (this._picker.current) {
                    this._picker.current.showPicker();
                }
            }
        }
        this.setValueInternal(e.currentTarget.textContent ?? "");
    }
}


function sendNotification(id: string) {
    var parameters = {};
    var entity = {};
    // @ts-ignore
    entity.id = id; entity.entityType = "email";
    // @ts-ignore
    parameters.entity = entity; parameters.IssueSend = true;

    var sendEmailRequest = {
        // @ts-ignore
        entity: parameters.entity, IssueSend: parameters.IssueSend,

        getMetadata: function () {
            return {
                boundParameter: "entity",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.email",
                        "structuralProperty": 5
                    },
                    "IssueSend": {
                        "typeName": "Edm.Boolean",
                        "structuralProperty": 1
                    }
                },
                operationType: 0,
                operationName: "SendEmail"
            };
        }
    };
    // @ts-ignore
    Xrm.WebApi.online.execute(sendEmailRequest).then(
        function success(result: { ok: any; responseText: string; }) {
            if (result.ok) {
                //var results = JSON.parse(result.responseText);
                console.log("Email Sent");
            }
        },
        function (error: { message: any; }) {
            // @ts-ignore
            Xrm.Utility.alertDialog(error.message);
        }
    );
}