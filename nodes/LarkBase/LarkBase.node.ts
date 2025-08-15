import {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	IDataObject,
	NodeOperationError,
	IHttpRequestOptions,
} from 'n8n-workflow';

export class LarkBase implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'LarkBase',
		name: 'larkBase',
		icon: 'file:larkbase.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Interact with LarkBase (Feishu Base) API',
		defaults: {
			name: 'LarkBase',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [],
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Record',
						value: 'record',
					},
				],
				default: 'record',
			},
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: {
					show: {
						resource: ['record'],
					},
				},
				options: [
					{
						name: 'Create',
						value: 'create',
						description: 'Create a record',
						action: 'Create a record',
					},
					{
						name: 'Delete',
						value: 'delete',
						description: 'Delete a record',
						action: 'Delete a record',
					},
					{
						name: 'Get',
						value: 'get',
						description: 'Get a record',
						action: 'Get a record',
					},
					{
						name: 'Get Many',
						value: 'getAll',
						description: 'Get many records',
						action: 'Get many records',
					},
					{
						name: 'Update',
						value: 'update',
						description: 'Update a record',
						action: 'Update a record',
					},
				],
				default: 'create',
			},
			// Tenant Access Token
			{
				displayName: 'Tenant Access Token',
				name: 'tenantAccessToken',
				type: 'string',
				required: true,
				default: '',
				description: 'The tenant access token for authentication',
			},
			// App Token
			{
				displayName: 'App Token',
				name: 'appToken',
				type: 'string',
				required: true,
				default: '',
				description: 'The app token of the LarkBase application',
			},
			// Table ID
			{
				displayName: 'Table ID',
				name: 'tableId',
				type: 'string',
				required: true,
				default: '',
				description: 'The ID of the table to operate on',
			},
			// Kind Type Records
			{
				displayName: 'Kind Type Records',
				name: 'kindTypeRecords',
				type: 'options',
				options: [
					{
						name: 'Map Each Columns',
						value: 'mapEachColumns',
					},
					{
						name: 'Send Raw Data',
						value: 'sendRawData',
					},
				],
				default: 'mapEachColumns',
				displayOptions: {
					show: {
						resource: ['record'],
						operation: ['create', 'update'],
					},
				},
			},
			// Mapping Column Mode
			{
				displayName: 'Mapping Column Mode',
				name: 'mappingColumnMode',
				type: 'options',
				options: [
					{
						name: 'Map Each Column Manually',
						value: 'mapEachColumnManually',
					},
					{
						name: 'Auto Map by Column Names',
						value: 'autoMapByColumnNames',
					},
				],
				default: 'mapEachColumnManually',
				displayOptions: {
					show: {
						resource: ['record'],
						operation: ['create', 'update'],
						kindTypeRecords: ['mapEachColumns'],
					},
				},
			},
			// Values to Send
			{
				displayName: 'Values to Send',
				name: 'valuesToSend',
				placeholder: 'Add Field',
				type: 'fixedCollection',
				typeOptions: {
					multipleValues: true,
					sortable: true,
				},
				displayOptions: {
					show: {
						resource: ['record'],
						operation: ['create', 'update'],
						kindTypeRecords: ['mapEachColumns'],
						mappingColumnMode: ['mapEachColumnManually'],
					},
				},
				default: {},
				options: [
					{
						name: 'fields',
						displayName: 'Field',
						values: [
							{
								displayName: 'Field Name or ID',
								name: 'fieldId',
								type: 'string',
								default: '',
								description: 'Name or ID of the field in LarkBase',
							},
							{
								displayName: 'Field Value',
								name: 'fieldValue',
								type: 'string',
								default: '',
								description: 'Value to set for the field',
							},
						],
					},
				],
			},
			// Record ID for operations that need it
			{
				displayName: 'Record ID',
				name: 'recordId',
				type: 'string',
				displayOptions: {
					show: {
						resource: ['record'],
						operation: ['get', 'update', 'delete'],
					},
				},
				default: '',
				description: 'The ID of the record',
			},
			// Limit for getAll
			{
				displayName: 'Limit',
				name: 'limit',
				type: 'number',
				displayOptions: {
					show: {
						resource: ['record'],
						operation: ['getAll'],
					},
				},
				typeOptions: {
					minValue: 1,
					maxValue: 500,
				},
				default: 50,
				description: 'Max number of results to return',
			},
			// Return All
			{
				displayName: 'Return All',
				name: 'returnAll',
				type: 'boolean',
				displayOptions: {
					show: {
						resource: ['record'],
						operation: ['getAll'],
					},
				},
				default: false,
				description: 'Whether to return all results or only up to the limit',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];
		
		const resource = this.getNodeParameter('resource', 0) as string;
		const operation = this.getNodeParameter('operation', 0) as string;

		for (let i = 0; i < items.length; i++) {
			try {
				const tenantAccessToken = this.getNodeParameter('tenantAccessToken', i) as string;
				const appToken = this.getNodeParameter('appToken', i) as string;
				const tableId = this.getNodeParameter('tableId', i) as string;

				if (resource === 'record') {
					let responseData: any;

					if (operation === 'create') {
						const kindTypeRecords = this.getNodeParameter('kindTypeRecords', i) as string;
						let fields: IDataObject = {};

						if (kindTypeRecords === 'mapEachColumns') {
							const mappingColumnMode = this.getNodeParameter('mappingColumnMode', i) as string;
							
							if (mappingColumnMode === 'mapEachColumnManually') {
								const valuesToSend = this.getNodeParameter('valuesToSend', i) as IDataObject;
								if (valuesToSend.fields) {
									const fieldsArray = valuesToSend.fields as Array<{fieldId: string, fieldValue: string}>;
									for (const field of fieldsArray) {
										fields[field.fieldId] = field.fieldValue;
									}
								}
							} else {
								// Auto map by column names
								fields = items[i].json;
							}
						} else {
							// Send raw data
							fields = items[i].json;
						}

						const options: IHttpRequestOptions = {
							method: 'POST',
							url: `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records`,
							headers: {
								'Authorization': `Bearer ${tenantAccessToken}`,
								'Content-Type': 'application/json',
							},
							body: { fields },
							json: true,
						};

						responseData = await this.helpers.httpRequest(options);

					} else if (operation === 'get') {
						const recordId = this.getNodeParameter('recordId', i) as string;
						
						const options: IHttpRequestOptions = {
							method: 'GET',
							url: `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records/${recordId}`,
							headers: {
								'Authorization': `Bearer ${tenantAccessToken}`,
								'Content-Type': 'application/json',
							},
							json: true,
						};

						responseData = await this.helpers.httpRequest(options);

					} else if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						const limit = this.getNodeParameter('limit', i) as number;

						const qs: IDataObject = {};
						if (!returnAll) {
							qs.page_size = limit;
						}

						const options: IHttpRequestOptions = {
							method: 'GET',
							url: `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records`,
							headers: {
								'Authorization': `Bearer ${tenantAccessToken}`,
								'Content-Type': 'application/json',
							},
							qs,
							json: true,
						};

						responseData = await this.helpers.httpRequest(options);

						// Handle pagination if returnAll is true
						if (returnAll && responseData.data?.has_more) {
							let allRecords = responseData.data.items || [];
							let pageToken = responseData.data.page_token;

							while (pageToken) {
								const nextOptions = { ...options };
								nextOptions.qs = { ...qs, page_token: pageToken };
								
								const nextResponse = await this.helpers.httpRequest(nextOptions);
								
								if (nextResponse.data?.items) {
									allRecords = allRecords.concat(nextResponse.data.items);
								}
								
								pageToken = nextResponse.data?.page_token;
								if (!nextResponse.data?.has_more) break;
							}

							responseData.data.items = allRecords;
						}

					} else if (operation === 'update') {
						const recordId = this.getNodeParameter('recordId', i) as string;
						const kindTypeRecords = this.getNodeParameter('kindTypeRecords', i) as string;
						let fields: IDataObject = {};

						if (kindTypeRecords === 'mapEachColumns') {
							const mappingColumnMode = this.getNodeParameter('mappingColumnMode', i) as string;
							
							if (mappingColumnMode === 'mapEachColumnManually') {
								const valuesToSend = this.getNodeParameter('valuesToSend', i) as IDataObject;
								if (valuesToSend.fields) {
									const fieldsArray = valuesToSend.fields as Array<{fieldId: string, fieldValue: string}>;
									for (const field of fieldsArray) {
										fields[field.fieldId] = field.fieldValue;
									}
								}
							} else {
								fields = items[i].json;
							}
						} else {
							fields = items[i].json;
						}

						const options: IHttpRequestOptions = {
							method: 'PUT',
							url: `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records/${recordId}`,
							headers: {
								'Authorization': `Bearer ${tenantAccessToken}`,
								'Content-Type': 'application/json',
							},
							body: { fields },
							json: true,
						};

						responseData = await this.helpers.httpRequest(options);

					} else if (operation === 'delete') {
						const recordId = this.getNodeParameter('recordId', i) as string;
						
						const options: IHttpRequestOptions = {
							method: 'DELETE',
							url: `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records/${recordId}`,
							headers: {
								'Authorization': `Bearer ${tenantAccessToken}`,
								'Content-Type': 'application/json',
							},
							json: true,
						};

						responseData = await this.helpers.httpRequest(options);
					}

					if (responseData.code !== 0) {
						throw new NodeOperationError(this.getNode(), `LarkBase API Error: ${responseData.msg}`, {
							itemIndex: i,
						});
					}

					const executionData = this.helpers.constructExecutionMetaData(
						[responseData],
						{ itemData: { item: i } },
					);
					returnData.push(...executionData);
				}

			} catch (error) {
				if (this.continueOnFail()) {
					const executionData = this.helpers.constructExecutionMetaData(
						[{ error: error.message }],
						{ itemData: { item: i } },
					);
					returnData.push(...executionData);
					continue;
				}
				throw error;
			}
		}

		return [returnData];
	}
}