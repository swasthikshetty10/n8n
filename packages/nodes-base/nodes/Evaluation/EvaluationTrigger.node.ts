import type {
	INodeType,
	INodeTypeDescription,
	IExecuteFunctions,
	INodeExecutionData,
	IDataObject,
} from 'n8n-workflow';
import { NodeConnectionTypes } from 'n8n-workflow';

import { loadOptions } from './methods';
import { document, sheet } from '../Google/Sheet/GoogleSheetsTrigger.node';
import { readFilter } from '../Google/Sheet/v2/actions/sheet/read.operation';
import { readSheet } from '../Google/Sheet/v2/actions/utils/readOperation';
import { authentication } from '../Google/Sheet/v2/actions/versionDescription';
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
			{
				displayName:
					'Pulls a test dataset from a Google Sheet. The workflow will run once for each row, in sequence. More info.', // TODO Change
				name: 'notice',
				type: 'notice',
				default: '',
			},
			authentication,
			document,
			sheet,
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
			readFilter,
		],
	};

	methods = { loadOptions };

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

			const sheetName = result.title;

			const maxRows = this.getNodeParameter('limitRows', 0)
				? (this.getNodeParameter('maxRows', 0) as string)
				: undefined;
			const rangeString = maxRows ? `${sheetName}!1:${maxRows}` : `${sheetName}!A:Z`;

			operationResult = await readSheet.call(
				this,
				googleSheet,
				sheetName,
				0,
				operationResult,
				5,
				[],
				rangeString,
			);
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
