import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'BootstrapCarouselWebPartStrings';
import BootstrapCarousel from './components/BootstrapCarousel';
import { IBootstrapCarouselProps } from './components/IBootstrapCarouselProps';

import { SPFI, SPFx, spfi } from '@pnp/sp';

export interface IBootstrapCarouselWebPartProps {
  description: string,
  maxSlides: number
}

export interface ISlide {
  title: string,
  bannerUrl: string,
  pageRelativeUrl: string
}

export default class BootstrapCarouselWebPart extends BaseClientSideWebPart<IBootstrapCarouselWebPartProps> {
  private sp: SPFI;

  protected async onInit(): Promise<void> {
    const sp = spfi().using(SPFx(this.context));

    this.sp = sp;
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IBootstrapCarouselProps> = (
      <BootstrapCarousel sp={this.sp} maxSlides={this.properties.maxSlides} />
    )

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
