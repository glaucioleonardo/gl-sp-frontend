import { SpCore } from '../setup/core-services-setup.service';
import { IListInfo, sp } from '@pnp/sp/presets/core';

import '@pnp/sp/fields';
import '@pnp/sp/lists';
import '@pnp/sp/webs';
import '@pnp/sp/views';

import { ISpCoreResult } from '../setup/core-services-setup.interface';
import { IListProperties } from './core-services-lists.interface';
import { CalendarType, DateTimeFieldFormatType, DateTimeFieldFriendlyFormatType, FieldTypes, IFieldInfo } from '@pnp/sp/fields';


class Core {

  fieldsToStringArray(fields: string[]): string {
    return fields.toString().replace('[', '').replace(']', '');
  }

  /**
   * Check if current list exists.
   * @param listName
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  async exists(listName: string, baseUrl?: string): Promise<boolean> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      const lists = await this.retrieve(listName, base);
      return lists.filter(x => x.Title === listName).length > 0;
    } catch (reason) {
      const error = SpCore.onError(reason)
      SpCore.showErrorLog(reason);
      throw new Error(error.code.toString())
    }
  }

  /**
   * Retrieve a unique list item based on fields, filter and ordering
   * @param listName
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  async retrieveSingle(listName: string, baseUrl?: string): Promise<IListInfo> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      return await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).get();
    } catch (reason) {
      const error = SpCore.onError(reason)
      SpCore.showErrorLog(reason);
      throw new Error(error.code.toString())
    }
  }

  /**
   * Retrieve a search list based on fields, filter and ordering
   * @param listName
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  async retrieve(listName: string, baseUrl?: string): Promise<IListInfo[]> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      return await sp.configure(SpCore.config, base).web.lists.get();
    } catch (reason) {
      const error = SpCore.onError(reason)
      SpCore.showErrorLog(reason);
      throw new Error(error.code.toString())
    }
  }

  /**
   * Move List to recycle bin. The user will be able to restore the information.
   * @param listName
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   */
  async recycle(listName: string, baseUrl?: string): Promise<ISpCoreResult> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      const exists: boolean = await this.exists(listName, base);

      if (exists) {
        await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).recycle()
        return {
          code: 200,
          description: 'Success!',
          message: 'The current list has been recycled.'
        };
      } else {
        return {
          code: 503,
          description: 'Internal Error!',
          message: 'The current list doesn\'t exist.'
        }
      }
    } catch (reason) {
      SpCore.showErrorLog(reason);
      return  {
        code: 500,
        description: 'Internal Error!',
        message: reason.message
      };
    }
  }

  /**
   * Recreates the lists recycling the old one.
   * @param listName
   * @param baseUrl (optional) In case it is necessary to gather data from another url.
   * @param fields (optional) Creates the columns
   * @param titleRequired (optional) Change the Title as a required field. Default is true
   * @param properties (optional) The list properties
   */
  async recreate(listName: string, baseUrl?: string, fields: Partial<IFieldInfo>[] = [], titleRequired: boolean = true, properties?: IListProperties): Promise<ISpCoreResult> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;
    const listProperties: IListProperties = properties != null ? properties : {
      AllowContentTypes: true,
      BaseTemplate: 100,
      BaseType: 0,
      ContentTypesEnabled: false,
      EnableAttachments: true,
      DocumentTemplateUrl: undefined,
      EnableVersioning: true,
      Description: ''
    }

    try {
      await this.recycle(listName, baseUrl);
      await sp.configure(SpCore.config, base).web.lists.add(listName, listProperties.Description, listProperties.BaseTemplate, false, listProperties);
      await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.getByTitle('Title').update({
        Required: titleRequired,
        __metadata: { type: 'SP.FieldText' }
      });
      await this.addFields(listName, fields, base);

      return {
        code: 200,
        description: 'Success!',
        message: 'The current list has been recycled.'
      };
    } catch (reason) {
      SpCore.showErrorLog(reason);
      return  {
        code: 500,
        description: 'Internal Error!',
        message: reason.message
      };
    }
  }

  /**
   * Add fields to the list
   */
  async addFields(listName: string, fields: Partial<IFieldInfo>[], baseUrl?: string): Promise<ISpCoreResult> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      for (const [i, field] of fields.entries()) {
        const title: string = field.Title == null ? `Column${i + 1}` : field.Title;
        const required: boolean = field.Required == null ? false : field.Required;
        const indexed: boolean = field.Indexed == null ? false : field.Indexed;
        const defaultBooleanValue: string = field.DefaultValue == null ? '0' : field.DefaultValue;

        switch (field.FieldTypeKind) {
          case FieldTypes.Text:
            await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.addText(title, undefined, {
              Required: required,
              Indexed: indexed
            });
            break;
          case FieldTypes.Note:
            await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.addMultilineText(title, 6, false, false, false, false,  {
              Required: required,
              Indexed: indexed
            });
            break;
          case FieldTypes.Boolean:
            await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.addBoolean(title, {
              Required: required,
              Indexed: indexed,
              DefaultValue: defaultBooleanValue
            });
            break;
          case FieldTypes.Number:
            await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.addNumber(title, undefined, undefined, {
              Required: required,
              Indexed: indexed,
              DefaultValue: defaultBooleanValue
            });
            break;
          case FieldTypes.DateTime:
            await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).fields.addDateTime(title, DateTimeFieldFormatType.DateOnly, CalendarType.Gregorian, DateTimeFieldFriendlyFormatType.Disabled, {
              Required: required,
              Indexed: indexed,
              DefaultValue: defaultBooleanValue
            });
            break;
        }

        await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).defaultView.fields.add(title);
      }

      return {
        code: 200,
        description: 'Success!',
        message: 'The current list has been recycled.'
      };
    } catch (reason) {
      SpCore.showErrorLog(reason);
      return  {
        code: 500,
        description: 'Internal Error!',
        message: reason.message
      };
    }
  }

  /**
   * Recover list from recycle bin
   */
  async rename(listName: string, name: string, overwrite: boolean = true, baseUrl?: string): Promise<ISpCoreResult> {
    const base = baseUrl == null ? SpCore.baseUrl : baseUrl;

    try {
      const exists: boolean = await this.exists(name, base);

      if (exists && overwrite) {
        await this.recycle(listName, base);
        await sp.configure(SpCore.config, base).web.lists.getByTitle(listName).update({
          Title: `_backup-${listName}`
        });

        return {
          code: 200,
          description: 'Success!',
          message: 'The current list has been renamed successfully.'
        };
      } else {
        if (exists && !overwrite) {
          return {
            code: 405,
            description: 'Not allowed!',
            message: 'The current list exists and was not allowed to overwrite it.'
          };
        } else {
          return {
            code: 404,
            message: 'Error making HttpClient request in queryable: [404] Not Found',
            description: 'List not found',
          };
        }
      }
    } catch (reason) {
      SpCore.showErrorLog(reason);
      throw new Error(reason)
    }
  }
}
export const ListsCore = new Core();

