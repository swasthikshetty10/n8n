import {
	IDataObject,
	IExecuteFunctions,
	INodeExecutionData,
	NodeOperationError,
	UserError,
} from 'n8n-workflow';
import { GoogleSheet } from '../../Google/Sheet/v2/helpers/GoogleSheet';
import {
	ROW_NUMBER,
	ValueInputOption,
	ValueRenderOption,
} from '../../Google/Sheet/v2/helpers/GoogleSheets.types';
import { cellFormatDefault } from '../../Google/Sheet/v2/helpers/GoogleSheets.utils';
import { ISheetUpdateData } from '../../Google/Sheet/v1/GoogleSheet';

export async function add(
	eFunc: IExecuteFunctions,
	sheet: GoogleSheet,
	range: string,
	sheetId: string,
): Promise<INodeExecutionData[]> {
	const items = eFunc.getInputData();
	const nodeVersion = 5;

	if (!items.length) return [];

	const options = eFunc.getNodeParameter('options', 0, {});
	const locationDefine = (options.locationDefine as IDataObject)?.values as IDataObject;

	let keyRowIndex = 1;
	if (locationDefine?.headerRow) {
		keyRowIndex = locationDefine.headerRow as number;
	}

	const sheetData = await sheet.getData(range, 'FORMATTED_VALUE');

	const inputData: IDataObject[] = [{ a: '1', b: '2' }];

	const valueInputMode = (options.cellFormat as ValueInputOption) || cellFormatDefault(nodeVersion);

	await sheet.appendEmptyRowsOrColumns(sheetId, 1, 0);

	// if sheetData is undefined it means that the sheet was empty
	// we did add row with column names in the first row (autoMapInputData)
	// to account for that length has to be 1 and we append data in the next row
	const lastRow = (sheetData ?? [{}]).length + 1;

	await sheet.appendSheetData({
		inputData,
		range,
		keyRowIndex,
		valueInputMode,
		lastRow,
	});

	return items.map((item, index) => {
		item.pairedItem = { item: index };
		return item;
	});
}

export async function update(
	eFunc: IExecuteFunctions,
	gsheet: GoogleSheet,
	sheetName: string,
): Promise<INodeExecutionData[]> {
	const items = eFunc.getInputData();
	const nodeVersion = eFunc.getNode().typeVersion;

	const range = `${sheetName}!A:Z`;

	const valueInputMode = eFunc.getNodeParameter(
		'options.cellFormat',
		0,
		cellFormatDefault(nodeVersion),
	) as ValueInputOption;

	const options = eFunc.getNodeParameter('options', 0, {});

	const valueRenderMode = (options.valueRenderMode || 'UNFORMATTED_VALUE') as ValueRenderOption;

	const locationDefineOptions = (options.locationDefine as IDataObject)?.values as IDataObject;

	let keyRowIndex = 0;
	let dataStartRowIndex = 1;

	if (locationDefineOptions) {
		if (locationDefineOptions.headerRow) {
			keyRowIndex = parseInt(locationDefineOptions.headerRow as string, 10) - 1;
		}
		if (locationDefineOptions.firstDataRow) {
			dataStartRowIndex = parseInt(locationDefineOptions.firstDataRow as string, 10) - 1;
		}
	}

	let columnNames: string[] = [];

	const sheetData = await gsheet.getData(sheetName, 'FORMATTED_VALUE');

	if (sheetData?.[keyRowIndex] === undefined) {
		throw new NodeOperationError(
			eFunc.getNode(),
			`Could not retrieve the column names from row ${keyRowIndex + 1}`,
		);
	}

	columnNames = sheetData[keyRowIndex];

	const newColumns = new Set<string>();

	const columnsToMatchOn: string[] = ['row_number'];

	const dataMode = 'defineBelow';

	// TODO: Add support for multiple columns to match on in the next overhaul
	const keyIndex = columnNames.indexOf(columnsToMatchOn[0]);

	//not used when updating row
	const columnValuesList = await gsheet.getColumnValues({
		range,
		keyIndex,
		dataStartRowIndex,
		valueRenderMode,
		sheetData,
	});

	const updateData: ISheetUpdateData[] = [];

	const mappedValues: IDataObject[] = [];

	const errorOnUnexpectedColumn = (key: string, i: number) => {
		if (!columnNames.includes(key)) {
			throw new NodeOperationError(eFunc.getNode(), 'Unexpected fields in node input', {
				itemIndex: i,
				description: `The input field '${key}' doesn't match any column in the Sheet. You can ignore this by changing the 'Handling extra data' field, which you can find under 'Options'.`,
			});
		}
	};

	for (let i = 0; i < items.length; i++) {
		const inputData: IDataObject[] = [];

		const valueToMatchOn = 'val';

		if (valueToMatchOn === '') {
			throw new NodeOperationError(
				eFunc.getNode(),
				"The 'Column to Match On' parameter is required",
				{
					itemIndex: i,
				},
			);
		}

		const mappingValues = { row_number: 1 };
		if (Object.keys(mappingValues).length === 0) {
			throw new NodeOperationError(
				eFunc.getNode(),
				"At least one value has to be added under 'Values to Send'",
			);
		}
		// Setting empty values to empty string so that they are not ignored by the API
		Object.keys(mappingValues).forEach((key) => {
			if (
				key === 'row_number' &&
				(mappingValues[key] === null || mappingValues[key] === undefined)
			) {
				throw new UserError(
					'Column to match on (row_number) is not defined. Since the field is used to determine the row to update, it needs to have a value set.',
				);
			}

			if (mappingValues[key] === undefined || mappingValues[key] === null) {
				mappingValues[key] = '';
			}
		});
		inputData.push(mappingValues);
		mappedValues.push(mappingValues);

		let preparedData;
		const columnNamesList = [columnNames.concat([...newColumns])];

		preparedData = gsheet.prepareDataForUpdatingByRowNumber(inputData, range, columnNamesList);

		updateData.push(...preparedData.updateData);
	}

	if (updateData.length) {
		await gsheet.batchUpdate(updateData, valueInputMode);
	}

	if (!updateData.length) {
		return [];
	}

	const returnData: INodeExecutionData[] = [];
	for (const [index, entry] of mappedValues.entries()) {
		returnData.push({
			json: entry,
			pairedItem: { item: index },
		});
	}
	return returnData;
}
