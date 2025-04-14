import type {
	INodeType,
	INodeTypeDescription,
	IExecuteFunctions,
	INodeExecutionData,
	IDataObject,
} from 'n8n-workflow';
import { NodeConnectionTypes } from 'n8n-workflow';

import { GOOGLE_DRIVE_FILE_URL_REGEX, GOOGLE_SHEETS_SHEET_URL_REGEX } from '../Google/constants';
import { readSheet } from '../Google/Sheet/v2/actions/utils/readOperation';
import { GoogleSheet } from '../Google/Sheet/v2/helpers/GoogleSheet';
import type { ResourceLocator } from '../Google/Sheet/v2/helpers/GoogleSheets.types';
import { getSpreadsheetId } from '../Google/Sheet/v2/helpers/GoogleSheets.utils';

export class EvaluationTrigger implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Evaluation Trigger',
		name: 'evaluationTrigger',
		group: ['trigger'],
		version: 1,
		description: 'Runs an evaluation',
		eventTriggerDescription: '',
		maxNodes: 1,
		defaults: {
			name: 'Evaluation Trigger',
		},

		inputs: [],
		outputs: [NodeConnectionTypes.Main],
		credentials: [
			{
				name: 'googleApi',
				required: true,
				displayOptions: {
					show: {
						authentication: ['serviceAccount'],
					},
				},
				testedBy: 'googleApiCredentialTest',
			},
			{
				name: 'googleSheetsOAuth2Api',
				required: true,
				displayOptions: {
					show: {
						authentication: ['oAuth2'],
					},
				},
			},
		],
		properties: [
			// trigger shared logic with GoogleSheets node, leaving this here for compatibility
			{
				displayName:
					'Pulls a test dataset from a Google Sheet. The workflow will run once for each row, in sequence. More info.', // TODO Change
				name: 'notice',
				type: 'notice',
				default: '',
			},
			{
				displayName: 'Authentication',
				name: 'authentication',
				type: 'options',
				options: [
					{
						name: 'Service Account',
						value: 'serviceAccount',
					},
					{
						// eslint-disable-next-line n8n-nodes-base/node-param-display-name-miscased
						name: 'OAuth2 (recommended)',
						value: 'oAuth2',
					},
				],
				default: 'oAuth2',
			},
			{
				displayName: 'Document',
				name: 'documentId',
				type: 'resourceLocator',
				default: { mode: 'list', value: '' },
				required: true,
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						typeOptions: {
							searchListMethod: 'spreadSheetsSearch',
							searchable: true,
						},
					},
					{
						displayName: 'By URL',
						name: 'url',
						type: 'string',
						extractValue: {
							type: 'regex',
							regex: GOOGLE_DRIVE_FILE_URL_REGEX,
						},
						validation: [
							{
								type: 'regex',
								properties: {
									regex: GOOGLE_DRIVE_FILE_URL_REGEX,
									errorMessage: 'Not a valid Google Drive File URL',
								},
							},
						],
					},
					{
						displayName: 'By ID',
						name: 'id',
						type: 'string',
						validation: [
							{
								type: 'regex',
								properties: {
									regex: '[a-zA-Z0-9\\-_]{2,}',
									errorMessage: 'Not a valid Google Drive File ID',
								},
							},
						],
						url: '=https://docs.google.com/spreadsheets/d/{{$value}}/edit',
					},
				],
			},
			{
				displayName: 'Sheet',
				name: 'sheetName',
				type: 'resourceLocator',
				default: { mode: 'list', value: '' },
				// default: '', //empty string set to progresivly reveal fields
				required: true,
				typeOptions: {
					loadOptionsDependsOn: ['documentId.value'],
				},
				modes: [
					{
						displayName: 'From List',
						name: 'list',
						type: 'list',
						typeOptions: {
							searchListMethod: 'sheetsSearch',
							searchable: false,
						},
					},
					{
						displayName: 'By URL',
						name: 'url',
						type: 'string',
						extractValue: {
							type: 'regex',
							regex: GOOGLE_SHEETS_SHEET_URL_REGEX,
						},
						validation: [
							{
								type: 'regex',
								properties: {
									regex: GOOGLE_SHEETS_SHEET_URL_REGEX,
									errorMessage: 'Not a valid Sheet URL',
								},
							},
						],
					},
					{
						displayName: 'By ID',
						name: 'id',
						type: 'string',
						validation: [
							{
								type: 'regex',
								properties: {
									regex: '((gid=)?[0-9]{1,})',
									errorMessage: 'Not a valid Sheet ID',
								},
							},
						],
					},
					{
						displayName: 'By Name',
						name: 'name',
						type: 'string',
						placeholder: 'Sheet1',
					},
				],
			},
			{
				displayName: 'Limit Rows',
				name: 'limitRows',
				type: 'boolean',
				default: false,
				noDataExpression: true,
				description: 'Whether to limit number of rows to process',
			},
			{
				displayName: 'Max Rows to Process',
				name: 'maxRows',
				type: 'string',
				default: '10',
				description: 'Maximum number of rows to process',
				noDataExpression: false,
				displayOptions: { show: { limitRows: [true] } },
			},
		],
	};

	async execute(this: IExecuteFunctions) {
		let operationResult: INodeExecutionData[] = [];

		try {
			const { mode, value } = this.getNodeParameter('documentId', 0) as IDataObject;
			const spreadsheetId = getSpreadsheetId(
				this.getNode(),
				mode as ResourceLocator,
				value as string,
			);

			const googleSheet = new GoogleSheet(spreadsheetId, this);

			let sheetId = '';
			let sheetName = '';

			const sheetWithinDocument = this.getNodeParameter('sheetName', 0, undefined, {
				extractValue: true,
			}) as string;
			const { mode: sheetMode } = this.getNodeParameter('sheetName', 0) as {
				mode: ResourceLocator;
			};

			const result = await googleSheet.spreadsheetGetSheet(
				this.getNode(),
				sheetMode,
				sheetWithinDocument,
			);
			sheetId = result.sheetId.toString();
			sheetName = result.title;

			const results = await readSheet.call(this, googleSheet, sheetName, 0, operationResult, 5, []);

			if (results?.length) {
				operationResult = operationResult.concat(results);
			}
		} catch (error) {
			if (this.continueOnFail()) {
				operationResult.push({ json: this.getInputData(0)[0].json, error });
			} else {
				throw error;
			}
		}

		return [operationResult];
	}
}
