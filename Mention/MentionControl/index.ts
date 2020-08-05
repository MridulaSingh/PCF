import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { IPersonaProps, IPersonaStyles } from "office-ui-fabric-react/lib/Persona";
import MentionEditorControl, { IMentionProps } from "./utilities/Mention";
import ReactDOM = require("react-dom");
import React = require("react");
import { people } from "@uifabric/example-data";

export class MentionControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private mentionEditor: MentionEditorControl;
	private _notifyOutputChanged: any;
//	private Xrm: any;
	
	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Add control initialization code
		
		this._context = context;
		this._container = container;
		this._notifyOutputChanged = notifyOutputChanged;
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		const field: string = context.parameters.field.raw || '';
		const allowedNumberOfCharacters = context.parameters.field.attributes?.MaxLength || 100;
		const sendEmail = context.parameters.sendEmail.raw;
		const fromUser = context.parameters.emailFromUserGuid.raw || '';
		const subject = context.parameters.emailSubject.raw || '';
		const description = context.parameters.emailContent.raw || '';
		const _props: IMentionProps = {
			context: this._context,
			people: this._retrieveSystemUsers(),
			value: field,
			formatNumber: (n: number): string => context.formatting.formatInteger(n),
			notifyOutputChanged: this._notifyOutputChanged,
			allowedNumberOfCharacters: allowedNumberOfCharacters,
			fromUser: fromUser,
			sendEmail: sendEmail,
			subject: subject,
			description: description
		}
		this.mentionEditor = ReactDOM.render(
			React.createElement(MentionEditorControl, _props),
			this._container
		);
		this.mentionEditor.setValue(context.parameters.field.raw);
		
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {
			field: this.mentionEditor.getValue() ?? undefined,
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

	private _retrieveSystemUsers(): IPersonaProps[] {
		let People: IPersonaProps[] = [];
		//@ts-ignore
		Xrm.WebApi.online.retrieveMultipleRecords("systemuser", "?$select=fullname,systemuserid,firstname,jobtitle,entityimage_url").then(
			function success(result: { entities: any[]; }) {
				result.entities.map((entity: any) => {
					People.push({
						//"styles":styleq,
						"text": entity.fullname, "secondaryText": entity.jobtitle,"primaryText":entity.firstname, "optionalText": entity.systemuserid,					
						"imageUrl":
							//@ts-ignore
							Xrm.Page.context.getClientUrl() + entity.entityimage_url
					});
				});

			});
		return People;
	}
        
}