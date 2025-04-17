import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpFourTilesWebPartStrings';
import SpFourTiles from './components/SpFourTiles';
import { ISpFourTilesProps, ITileProps } from './components/ISpFourTilesProps';

export interface ISpFourTilesWebPartProps {
  tile1: ITileProps;
  tile2: ITileProps;
  tile3: ITileProps;
  tile4: ITileProps;
  [key: string]: any; // Add index signature to allow string indexing
}

export default class SpFourTilesWebPart extends BaseClientSideWebPart<ISpFourTilesWebPartProps> {

  private _isDarkTheme: boolean = false;

  protected onInit(): Promise<void> {
    // Initialize default property values if they don't exist
    if (!this.properties.tile1) {
      this.properties.tile1 = {
        header: 'Cloud Entry',
        text: 'Architecture reviews, POC preparation, and hands-on workshops'
      };
    }
    
    if (!this.properties.tile2) {
      this.properties.tile2 = {
        header: 'Cloud Operations',
        text: 'Managed DevOps services for your applications, infrastructure, and services'
      };
    }
    
    if (!this.properties.tile3) {
      this.properties.tile3 = {
        header: 'Cloud Adoption',
        text: 'Application assessment, migration, and optimization for cost and performance'
      };
    }
    
    if (!this.properties.tile4) {
      this.properties.tile4 = {
        header: 'Cloud Solutions',
        text: 'Cloud-native applications using the power of serverless connected services'
      };
    }
    
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<ISpFourTilesProps> = React.createElement(
      SpFourTiles,
      {
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        tile1: this.properties.tile1,
        tile2: this.properties.tile2,
        tile3: this.properties.tile3,
        tile4: this.properties.tile4,
        displayMode: this.displayMode,
        updateProperty: (propertyPath: string, newValue: string) => {
          // Split the property path into parts
          const pathParts = propertyPath.split('.');
          
          // Handle nested property updates
          if (pathParts.length > 1) {
            const parentProp = pathParts[0] as keyof ISpFourTilesWebPartProps;
            const childProp = pathParts[1];
            
            // Create a new object with the updated property
            const updatedObject = {
              ...this.properties[parentProp],
              [childProp]: newValue
            };
            
            // Update the property
            this.properties[parentProp] = updatedObject;
          } else {
            // Update simple property
            const propName = propertyPath as keyof ISpFourTilesWebPartProps;
            this.properties[propName] = newValue;
          }
          
          // Trigger re-render
          this.render();
        }
      }
    );

    ReactDom.render(element, this.domElement);
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
              groupName: 'Tile 1',
              groupFields: [
                PropertyPaneTextField('tile1.header', {
                  label: 'Header'
                }),
                PropertyPaneTextField('tile1.text', {
                  label: 'Text',
                  multiline: true
                })
              ]
            },
            {
              groupName: 'Tile 2',
              groupFields: [
                PropertyPaneTextField('tile2.header', {
                  label: 'Header'
                }),
                PropertyPaneTextField('tile2.text', {
                  label: 'Text',
                  multiline: true
                })
              ]
            },
            {
              groupName: 'Tile 3',
              groupFields: [
                PropertyPaneTextField('tile3.header', {
                  label: 'Header'
                }),
                PropertyPaneTextField('tile3.text', {
                  label: 'Text',
                  multiline: true
                })
              ]
            },
            {
              groupName: 'Tile 4',
              groupFields: [
                PropertyPaneTextField('tile4.header', {
                  label: 'Header'
                }),
                PropertyPaneTextField('tile4.text', {
                  label: 'Text',
                  multiline: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
