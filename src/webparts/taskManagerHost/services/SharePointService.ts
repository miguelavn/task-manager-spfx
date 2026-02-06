import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';
import type { ITask } from '../components/ITask';

interface IListInfo {
  ListItemEntityTypeFullName: string;
}

interface ISPAuthor {
  Title?: string;
}

interface ISPTaskItem {
  Id?: number;
  Title?: string;
  Responsable?: string;
  Estado?: string;
  Author?: ISPAuthor;
  Created?: string;
  Modified?: string;
}

interface ISPItemsResponse<TItem> {
  value?: TItem[];
}

interface ISPODataVerboseResponse<T> {
  d?: T;
}

interface ISPODataVerboseResultsResponse<TItem> {
  d?: {
    results?: TItem[];
  };
}

export class SharePointService {
  private readonly _context: WebPartContext;
  private readonly _listTitle: string = 'TaskManager';
  private readonly _webUrl: string;
  private _listItemEntityTypeFullName: string | undefined;

  private async _getWithAcceptFallback(url: string): Promise<SPHttpClientResponse> {
    // Orden pensado para máxima compatibilidad:
    // 1) JSON Light sin metadata (ideal)
    // 2) application/json (algunos tenants lo exigen)
    // 3) verbose (fallback clásico)
    const acceptHeaders: string[] = [
      'application/json;odata=nometadata',
      'application/json',
      'application/json;odata=verbose'
    ];

    let lastResponse: SPHttpClientResponse | undefined;
    for (const accept of acceptHeaders) {
      const response = await this._context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': accept } }
      );
      lastResponse = response;
      if (response.status !== 406) {
        return response;
      }
    }

    // Si llegó aquí, todas devolvieron 406
    return lastResponse as SPHttpClientResponse;
  }

  private async _postWithAcceptFallback(
    url: string,
    bodyNoMetadata: object,
    bodyVerbose?: object,
    additionalHeaders?: Record<string, string>
  ): Promise<SPHttpClientResponse> {
    type Attempt = { accept: string; contentType: string; body: object };
    const attempts: Attempt[] = [
      {
        accept: 'application/json;odata=nometadata',
        contentType: 'application/json;odata=nometadata',
        body: bodyNoMetadata
      },
      {
        accept: 'application/json',
        contentType: 'application/json',
        body: bodyNoMetadata
      }
    ];

    if (bodyVerbose) {
      attempts.push({
        accept: 'application/json;odata=verbose',
        contentType: 'application/json;odata=verbose',
        body: bodyVerbose
      });
    }

    let lastResponse: SPHttpClientResponse | undefined;
    for (const attempt of attempts) {
      const response = await this._context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': attempt.accept,
            'Content-Type': attempt.contentType,
            ...(additionalHeaders || {})
          },
          body: JSON.stringify(attempt.body)
        }
      );
      lastResponse = response;
      if (response.status !== 406) {
        return response;
      }
    }

    return lastResponse as SPHttpClientResponse;
  }

  public constructor(context: WebPartContext) {
    this._context = context;
    this._webUrl = context.pageContext.web.absoluteUrl;
  }

  private _escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }

  private async _tryGetListInfo(): Promise<IListInfo | undefined> {
    const listTitleEscaped = this._escapeODataString(this._listTitle);
    const url = `${this._webUrl}/_api/web/lists/getbytitle('${listTitleEscaped}')?$select=ListItemEntityTypeFullName`;
    const response: SPHttpClientResponse = await this._getWithAcceptFallback(url);

    if (!response.ok) {
      if (response.status === 404) {
        return undefined;
      }

      const text = await response.text();
      throw new Error(
        `No se pudo acceder a la lista '${this._listTitle}': ${response.status} ${response.statusText} ${text}`
      );
    }

    const json = await response.json();
    const verbose = (json as ISPODataVerboseResponse<IListInfo>).d;
    if (verbose?.ListItemEntityTypeFullName) {
      return verbose;
    }

    return json as IListInfo;
  }

  private async _fieldExists(internalName: string): Promise<boolean> {
    const listTitleEscaped = this._escapeODataString(this._listTitle);
    const internalNameEscaped = this._escapeODataString(internalName);
    const url = `${this._webUrl}/_api/web/lists/getbytitle('${listTitleEscaped}')/fields?$select=InternalName&$filter=InternalName eq '${internalNameEscaped}'`;
    const response: SPHttpClientResponse = await this._getWithAcceptFallback(url);

    if (!response.ok) {
      const text = await response.text();
      throw new Error(
        `No se pudieron verificar los campos de la lista '${this._listTitle}': ${response.status} ${response.statusText} ${text}`
      );
    }

    const json = await response.json();
    const nonVerbose = json as ISPItemsResponse<{ InternalName?: string }>;
    const verbose = (json as ISPODataVerboseResultsResponse<{ InternalName?: string }>).d?.results;
    const fields = nonVerbose.value || verbose || [];
    return fields.some((f) => (f.InternalName || '').toLowerCase() === internalName.toLowerCase());
  }

  public async ensureListExists(): Promise<void> {
    const listInfo = await this._tryGetListInfo();
    if (listInfo?.ListItemEntityTypeFullName) {
      this._listItemEntityTypeFullName = listInfo.ListItemEntityTypeFullName;
      await this._ensureFields();
      return;
    }

    await this._createTaskList();

    const listInfoAfter = await this._tryGetListInfo();
    if (listInfoAfter?.ListItemEntityTypeFullName) {
      this._listItemEntityTypeFullName = listInfoAfter.ListItemEntityTypeFullName;
      await this._ensureFields();
      return;
    }

    throw new Error(
      `Se intentó crear la lista '${this._listTitle}', pero no se pudieron obtener sus metadatos. Reintenta o valida permisos.`
    );
  }

  private async _createTaskList(): Promise<void> {
    const createUrl = `${this._webUrl}/_api/web/lists`;
    const bodyNoMetadata = {
      'BaseTemplate': 100,
      'Title': this._listTitle,
      'Description': 'Lista de tareas del Task Manager',
      'ContentTypesEnabled': true
    };

    const bodyVerbose = {
      '__metadata': { 'type': 'SP.List' },
      ...bodyNoMetadata
    };

    const response: SPHttpClientResponse = await this._postWithAcceptFallback(
      createUrl,
      bodyNoMetadata,
      bodyVerbose
    );

    if (!response.ok && response.status !== 409) {
      const text = await response.text();
      if (response.status === 401 || response.status === 403) {
        throw new Error(
          `No tienes permisos para crear la lista '${this._listTitle}' en este sitio (${response.status} ${response.statusText}). ` +
          `Pídele a un propietario/administrador del sitio que cree la lista y sus campos (Responsable, Estado) o que te otorgue permisos para administrar listas. ` +
          `${text}`
        );
      }

      throw new Error(
        `No se pudo crear la lista '${this._listTitle}': ${response.status} ${response.statusText} ${text}`
      );
    }
  }

  private async _ensureFields(): Promise<void> {
    if (!(await this._fieldExists('Responsable'))) {
      await this._ensureFieldXml(this._getResponsableFieldXml());
    }

    if (!(await this._fieldExists('Estado'))) {
      await this._ensureFieldXml(this._getEstadoFieldXml());
    }
  }

  private async _ensureFieldXml(schemaXml: string): Promise<void> {
    const listTitleEscaped = this._escapeODataString(this._listTitle);
    const url = `${this._webUrl}/_api/web/lists/getbytitle('${listTitleEscaped}')/fields/createfieldasxml`;
    const bodyNoMetadata = {
      'parameters': {
        'SchemaXml': schemaXml
      }
    };

    const bodyVerbose = {
      'parameters': {
        '__metadata': { 'type': 'SP.XmlSchemaFieldCreationInformation' },
        'SchemaXml': schemaXml
      }
    };

    const response: SPHttpClientResponse = await this._postWithAcceptFallback(
      url,
      bodyNoMetadata,
      bodyVerbose
    );

    if (!response.ok) {
      const text = await response.text();
      if (response.status === 401 || response.status === 403) {
        throw new Error(
          `No tienes permisos para crear campos en la lista '${this._listTitle}' (${response.status} ${response.statusText}). ` +
          `Un propietario/administrador debe crear los campos requeridos: Responsable (Texto) y Estado (Elección). ` +
          `${text}`
        );
      }

      throw new Error(
        `No se pudo crear un campo en la lista '${this._listTitle}': ${response.status} ${response.statusText} ${text}`
      );
    }
  }

  private _getResponsableFieldXml(): string {
    return "<Field DisplayName='Responsable' Name='Responsable' StaticName='Responsable' Type='Text' Group='Task Manager' />";
  }

  private _getEstadoFieldXml(): string {
    return (
      "<Field DisplayName='Estado' Name='Estado' StaticName='Estado' Type='Choice' Format='Dropdown' Group='Task Manager'>" +
      '<Default>Pendiente</Default>' +
      '<CHOICES>' +
      '<CHOICE>Pendiente</CHOICE>' +
      '<CHOICE>En Progreso</CHOICE>' +
      '<CHOICE>Completada</CHOICE>' +
      '<CHOICE>Cancelada</CHOICE>' +
      '</CHOICES>' +
      '</Field>'
    );
  }

  public async getTasks(): Promise<ITask[]> {
    await this.ensureListExists();

    const listTitleEscaped = this._escapeODataString(this._listTitle);
    const url = `${this._webUrl}/_api/web/lists/getbytitle('${listTitleEscaped}')/items` +
      "?$select=Id,Title,Responsable,Estado,Author/Title,Created,Modified" +
      '&$expand=Author' +
      '&$orderby=Id desc';

    const response: SPHttpClientResponse = await this._getWithAcceptFallback(url);

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`No se pudieron cargar las tareas: ${response.status} ${response.statusText} ${text}`);
    }

    const json = await response.json();
    const nonVerbose = json as ISPItemsResponse<ISPTaskItem>;
    const verbose = (json as ISPODataVerboseResultsResponse<ISPTaskItem>).d?.results;
    const items: ISPTaskItem[] = nonVerbose.value || verbose || [];

    return items.map((item: ISPTaskItem): ITask => {
      const authorTitle = item.Author?.Title || '';
      return {
        ID: item.Id,
        Title: item.Title || '',
        Responsable: item.Responsable || '',
        Estado: item.Estado || 'Pendiente',
        Autor: authorTitle,
        Creado: item.Created || '',
        Modificado: item.Modified || ''
      };
    });
  }

  public async createTask(task: ITask): Promise<void> {
    await this.ensureListExists();

    const listTitleEscaped = this._escapeODataString(this._listTitle);
    const url = `${this._webUrl}/_api/web/lists/getbytitle('${listTitleEscaped}')/items`;

    const bodyNoMetadata = {
      'Title': task.Title,
      'Responsable': task.Responsable,
      'Estado': task.Estado || 'Pendiente'
    };

    const entityType = this._listItemEntityTypeFullName || 'SP.ListItem';
    const bodyVerbose = {
      '__metadata': { 'type': entityType },
      ...bodyNoMetadata
    };

    const response: SPHttpClientResponse = await this._postWithAcceptFallback(
      url,
      bodyNoMetadata,
      bodyVerbose
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`No se pudo crear la tarea: ${response.status} ${response.statusText} ${text}`);
    }
  }

  public async updateTask(taskId: number, task: ITask): Promise<void> {
    await this.ensureListExists();

    const listTitleEscaped = this._escapeODataString(this._listTitle);
    const url = `${this._webUrl}/_api/web/lists/getbytitle('${listTitleEscaped}')/items(${taskId})`;

    const bodyNoMetadata = {
      'Title': task.Title,
      'Responsable': task.Responsable,
      'Estado': task.Estado
    };

    const entityType = this._listItemEntityTypeFullName || 'SP.ListItem';
    const bodyVerbose = {
      '__metadata': { 'type': entityType },
      ...bodyNoMetadata
    };

    const response: SPHttpClientResponse = await this._postWithAcceptFallback(
      url,
      bodyNoMetadata,
      bodyVerbose,
      {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      }
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`No se pudo actualizar la tarea: ${response.status} ${response.statusText} ${text}`);
    }
  }

  public async deleteTask(taskId: number): Promise<void> {
    await this.ensureListExists();

    const listTitleEscaped = this._escapeODataString(this._listTitle);
    const url = `${this._webUrl}/_api/web/lists/getbytitle('${listTitleEscaped}')/items(${taskId})`;
    const response: SPHttpClientResponse = await this._postWithAcceptFallback(
      url,
      {},
      {},
      {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    );

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`No se pudo eliminar la tarea: ${response.status} ${response.statusText} ${text}`);
    }
  }
}
