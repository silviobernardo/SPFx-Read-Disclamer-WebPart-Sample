import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeProvider, ThemeChangedEventArgs } from '@microsoft/sp-component-base';

import * as strings from 'ReadDisclamerWebPartStrings';
import ReadDisclamer from './components/ReadDisclamer';
import { IReadDisclamerProps, IReadDisclamerWebPartProps } from './components/IReadDisclamerProps';
import { sp } from "@pnp/sp/presets/all";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';

export default class ReadDisclamerWebPart extends BaseClientSideWebPart<IReadDisclamerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;
  private _currentUser: ISiteUserInfo | undefined;

  public render(): void {
    const configured: boolean = this.properties.storageList && this.properties.documentTitle
      ? this.properties.storageList !== '' && this.properties.documentTitle !== ''
      : false;

    const element: React.ReactElement<IReadDisclamerProps> = React.createElement(
      ReadDisclamer,
      {
        documentTitle: this.properties.documentTitle,
        storageList: this.properties.storageList,
        acknowledgementLabel: this.properties.acknowledgementLabel,
        acknowledgementMessage: this.properties.acknowledgementMessage,
        readMessage: this.properties.readMessage,
        themeVariant: this._themeVariant,
        configured,
        context: this.context,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        currentUser: this._currentUser,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    sp.setup({
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      spfxContext: this.context as any
    });

    // Get current user
    this._currentUser = await sp.web.currentUser();  // OR //this.context.pageContext.legacyPageContext.userId

    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
  }

  /**
 * Update the current theme variant reference and re-render.
 *
 * @param args The new theme
 */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('storageList', {
                  label: strings.StorageListLabel,
                  selectedList: this.properties.storageList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  multiSelect: false,
                  baseTemplate: 100 // Remove this property to get all Libraries and Lists
                }),
                PropertyPaneTextField('documentTitle', {
                  label: strings.DocumentTitleLabel,
                }),
                PropertyPaneTextField('acknowledgementLabel', {
                  label: strings.AcknowledgmentLabelLabel,
                }),
                PropertyPaneTextField('acknowledgementMessage', {
                  label: strings.AcknowledgmentMessageLabel,
                }),
                PropertyPaneTextField('readMessage', {
                  label: strings.ReadMessageLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
