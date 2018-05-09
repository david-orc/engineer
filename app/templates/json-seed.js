const seed = {
  name: '_Tile',
  url: '/',
  siteColumns: [
    {
      name: 'Test_Title',
      type: 'Text'
    },
    {
      name: 'Test_Icon',
      type: 'Text'
    },
    {
      name: 'Test_Label',
      type: 'Text'
    },
    {
      name: 'Test_Link',
      type: 'Url',
    },
  ],
  contentTypes: [
    {
      name: 'Test_Tile',
      siteColumns: ['Test_Title', 'Test_Icon', 'Test_Label', 'Test_Link'],
    },
  ],
  lists: [
    {
      name: 'Test_Tiles',
      contentTypes: ['Test_Tile'],
    },
  ]
};

const FIELD_TYPE = {
  // Add Boolean Field
  Boolean: {
    Type: 'Boolean',
    DefaultValue: '0'
  },

  // Add Choice field
  Choice: {
    Type: 'Choice',
    Choices: ['One', 'Two', 'Three']
  },

  // Add Currency field
  Currency: {
    Type: 'Currency',
    CurrencyLocaleId: 1033,
  },

  // Add DateTime field
  DateTime: {
    Type: 'DateTime',
    DisplayFormat: 1,
    DateTimeCalendarType: 1,
    FriendlyDisplayFormat: 0,
  },

  // Add MultiChoice field
  MultiChoice: {
    Type: 'MultiChoice',
    Choices: ['One', 'Two', 'Three'],
  },

  // Add MultiLineText field
  MultiLineText: {
    Type: 'MultiLineText',
    NumberOfLines: 8,
    RichText: true,
    AllowHyperlink: true,
    RestrictedMode: false,
    AppendOnly: false,
  },

  // Add Number field
  Number: {
    Type: 'Number',
    MinimumValue: 0,
    MaximumValue: 100,
  },

  // Add Text field
  Text: {
    Type: 'Text',
    MaxLength: 255,
  },

  // Add Url field
  Url: {
    Type: 'Url',
    DisplayFormat: 1,
  },

  // Add User field
  User: {
    Type: 'User',
    SelectionGroup: 8,
    SelectionMode: 0,
  },

  // Add Calculated field
  Calculated: {
    Type: 'Calculated',
    Formula: '=[NewNumberField]+10',
    OutputType: 'Number',
  },
};

function createFieldXML(attributeObject) {
  return `
      <Field
        ${Object.keys(attributeObject)
    .map(attrKey => attrKey + '="' + attributeObject[attrKey] + '"')
    .join(' ')}
      </Field>
    `;
}
const XML_FIELD_TYPE = {
  HTML: {
    Type: 'HTML',
    Name: 'TestHTMLField',
    StaticName: 'TestHTMLField',
    DisplayName: 'TestHTMLField',
    Group: '_Test Site Fields',
    RichText: 'TRUE',
    RichTextMode: 'FullHtml',
    Required: 'FALSE',
    SourceID: 'http://schemas.microsoft.com/sharepoint/v3'
  },

  // Image field
  Image: {
    Type: 'Image',
    Name: 'TestImageField',
    StaticName: 'TestImageField',
    DisplayName: 'TestImageField',
    Group: '_Test Site Fields',
    RichText: 'TRUE',
    RichTextMode: 'FullHtml',
    Required: 'FALSE',
    SourceID: 'http://schemas.microsoft.com/sharepoint/v3'
  },

  // Link field
  Link: {
    Type: 'Link',
    Name: 'TestLinkField',
    StaticName: 'TestLinkField',
    DisplayName: 'TestLinkField',
    Group: '_Test Site Fields',
    RichText: 'TRUE',
    RichTextMode: 'ThemeHtml',
    Required: 'FALSE',
    SourceID: 'http://schemas.microsoft.com/sharepoint/v3'
  },

  // SummaryLinks field
  SummaryLinks: {
    Type: 'SummaryLinks',
    Name: 'TestSummaryLinksField',
    StaticName: 'TestSummaryLinksField',
    DisplayName: 'TestSummaryLinksField',
    Group: '_Test Site Fields',
    RichText: 'TRUE',
    RichTextMode: 'FullHtml',
    Required: 'FALSE',
    SourceID: 'http://schemas.microsoft.com/sharepoint/v3'
  }
};

module.exports = {
  up(engineer) {
    const site = engineer.getWeb(seed.url);

    seed.siteColumns.forEach((sc) => {
      if (FIELD_TYPE[sc.type]) {        
        site.fields.add(Object.assign(
          { Title: sc.name, Group: seed.name },
          FIELD_TYPE[sc.type]
        ));
      } else if (XML_FIELD_TYPE[sc.type]) {
        site.fields.addXML(createFieldXML(Object.assign(
          { Name: sc.name, StaticName: sc.name, DisplayName: sc.name, Group: seed.name },
          XML_FIELD_TYPE[sc.type]
        )));
      } else {
        console.log(`No support for type:${sc.type}`);
      }
    });

    seed.contentTypes.forEach((ct) => {
      site.contentTypes.add({
        Name: ct.name,
        Group: seed.name,
        ParentContentTypeId: '0x01',
      });
      ct.siteColumns.forEach(sc =>
        site.contentTypes.getByName(ct.name).fieldLinks.add(sc)
      );
    });

    seed.lists.forEach((list) => {
      site.lists.add({
        Title: list.name,
        BaseTemplate: 100,
        ContentTypesEnabled: true,
      });
      list.contentTypes.forEach(ct =>
        site.lists.getByTitle(list.name).contentTypes.add(ct)
      );
    });
  },

  // down(engineer) {}
};