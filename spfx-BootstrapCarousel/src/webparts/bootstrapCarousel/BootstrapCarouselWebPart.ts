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
//import { SPFx, graphfi } from '@pnp/graph';
//import "@pnp/graph/lists";
//import "@pnp/graph/sites";

import { SPFx, spfi } from '@pnp/sp';
import "@pnp/sp/sites";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages";
import "@pnp/sp/files/folder";
import "@pnp/sp/folders";
import { IItem } from '@pnp/sp/items';

export interface IBootstrapCarouselWebPartProps {
  description: string;
}

export interface ISlide {
  title: string,
  bannerUrl: string
}

export default class BootstrapCarouselWebPart extends BaseClientSideWebPart<IBootstrapCarouselWebPartProps> {
  slides;

  protected async onInit(): Promise<void> {
    //const graph = graphfi().using(SPFx(this.context));
    const sp = spfi().using(SPFx(this.context));

    const sitePagesFiles = await sp.web.folders.getByUrl("SitePages").files();
    const sitePagesPromises = sitePagesFiles.map(v => sp.web.loadClientsidePage(v.ServerRelativeUrl));
    const sitePages = await Promise.all(sitePagesPromises);
    const sitePagesWithBanner = sitePages.filter(v => v.bannerImageUrl && v.bannerImageUrl !== '');
    
    this.slides = sitePagesWithBanner.map(v => ({ bannerUrl: v.bannerImageUrl, title: v.title }));
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IBootstrapCarouselProps> = React.createElement(
      BootstrapCarousel,
      {
        description: this.properties.description,
        slides: this.slides
      }
    );

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
