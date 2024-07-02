import * as React from 'react';
import { IBootstrapCarouselProps } from './IBootstrapCarouselProps';
import { IBootstrapCarouselState } from './IBootstrapCarouselState';

import { Carousel } from 'react-bootstrap';
import './BootstrapCarousel.module.scss';
import 'bootstrap/dist/css/bootstrap.min.css';

import '@pnp/sp/webs';
import '@pnp/sp/folders';
import '@pnp/sp/files';
import '@pnp/sp/clientside-pages';
import * as strings from 'BootstrapCarouselWebPartStrings';
import { ISlide } from '../BootstrapCarouselWebPart';

export default class BootstrapCarousel extends React.Component<IBootstrapCarouselProps, IBootstrapCarouselState> {

  constructor(props: IBootstrapCarouselProps) {
    super(props);
    this.state = {
      index: 0
    };

    const { sp } = this.props;

    sp.web.folders
      .getByUrl("SitePages")
      .files
      .orderBy('TimeCreated', false)
      .top(this.props.maxSlides)
      .select('ServerRelativeUrl')()
      .then(async pagesFiles => {
        const sitePages = await Promise.all(pagesFiles.map(v => sp.web.loadClientsidePage(v.ServerRelativeUrl)));
        const sitePagesWithBanner = sitePages.filter(v => v.bannerImageUrl && v.bannerImageUrl !== '');
        const slides = sitePagesWithBanner
          .map<ISlide>((v, i) => ({ bannerUrl: v.bannerImageUrl, title: v.title, pageRelativeUrl: pagesFiles[i].ServerRelativeUrl }))

        this.setState({ slides });
      })
      .catch(reason => this.setState({ errorMessage: reason }));

    this.handleSelect = this.handleSelect.bind(this);
  }

  public render(): React.ReactElement<IBootstrapCarouselProps> {
    const { slides } = this.state;
    return (
      <Carousel
        activeIndex={this.state.index}
        onSelect={this.handleSelect}
        indicatorLabels={this.state.slides && this.state.slides.map(v => v.title)}
      >
        {(slides && slides.length !== 0) ? this.state.slides
          .map((v, i) =>
            <Carousel.Item key={i}>
              <div style={{ position: 'relative', cursor: 'pointer' }} onClick={() => window.open(v.pageRelativeUrl, '_self')}>
                <img className='d-block w-100' style={{ aspectRatio: "2/1", objectFit: 'cover' }} src={v.bannerUrl} />
                <div style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', background: "0% 100% / auto 30% repeat-x linear-gradient(to top, #000000ff 0%, #00000000 100%)" }} />
              </div>
              <Carousel.Caption><h3>{v.title}</h3></Carousel.Caption>
            </Carousel.Item>
          )
          :
          <Carousel.Item>
            <div style={{ width: '100%', aspectRatio: '2/1', backgroundColor: 'gray' }} />
            <Carousel.Caption><h3>{this.state.errorMessage || (slides ? strings.NoNews : strings.Loading)}</h3></Carousel.Caption>
          </Carousel.Item>
        }
      </Carousel>
    );
  }

  private handleSelect(eventKey: number) {
    this.setState({
      index: eventKey
    });
  }

}
