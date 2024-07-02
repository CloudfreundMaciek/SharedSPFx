import { ISlide } from "../BootstrapCarouselWebPart";

export interface IBootstrapCarouselState {
  index: number,
  slides?: ISlide[],
  errorMessage?: string,
}
