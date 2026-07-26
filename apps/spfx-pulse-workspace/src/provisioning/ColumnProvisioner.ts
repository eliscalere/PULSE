import { SPFI } from '@pnp/sp';

import { SharePointFieldSchema } from '@/types/Schema';

export class ColumnProvisioner {
  public constructor(
    private readonly sp: SPFI,
    private readonly resolveListTitle: (listKey: string) => string
  ) {}

  public async ensureColumns(listTitle: string, fields: readonly SharePointFieldSchema[]): Promise<void> {
    const list = this.sp.web.lists.getByTitle(listTitle);
    const existingFields = await list.fields.select('InternalName')();
    const existingNames = new Set(existingFields.map((field) => field.InternalName));

    for (const field of fields) {
      if (existingNames.has(field.internalName)) {
        continue;
      }

      const xml = await this.buildFieldXml(field);
      await list.fields.createFieldAsXml(xml);
    }
  }

  private async buildFieldXml(field: SharePointFieldSchema): Promise<string> {
    const required = field.required ? 'TRUE' : 'FALSE';
    const indexed = field.indexed ? 'TRUE' : 'FALSE';
    const multi = field.multiValue ? 'TRUE' : 'FALSE';

    if (field.type === 'Choice' && field.choices) {
      const choices = field.choices.map((choice) => `<CHOICE>${choice.value}</CHOICE>`).join('');
      return `<Field Type="Choice" DisplayName="${field.displayName}" Name="${field.internalName}" Required="${required}" Indexed="${indexed}"><CHOICES>${choices}</CHOICES></Field>`;
    }

    if (field.type === 'Lookup' && field.lookup) {
      const lookupTitle = this.resolveListTitle(field.lookup.listKey);
      const lookupList = await this.sp.web.lists.getByTitle(lookupTitle).select('Id')();
      return `<Field Type="Lookup" DisplayName="${field.displayName}" Name="${field.internalName}" List="{${lookupList.Id}}" ShowField="${field.lookup.showField}" Mult="${multi}" Required="${required}" />`;
    }

    if (field.type === 'User') {
      return `<Field Type="User" DisplayName="${field.displayName}" Name="${field.internalName}" UserSelectionMode="PeopleOnly" Mult="${multi}" Required="${required}" />`;
    }

    return `<Field Type="${field.type}" DisplayName="${field.displayName}" Name="${field.internalName}" Required="${required}" Indexed="${indexed}" />`;
  }
}
