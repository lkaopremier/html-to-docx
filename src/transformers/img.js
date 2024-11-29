import { imageBase64ToBuffer, splitMeasure } from "../helper.js";
import Transformer from "./transform.js";

class ImgTransformer extends Transformer {
  async transform() {
    const image = await this.getImage();

    if (!image) return [];

    const style = this.computeStyle(image);

    return [
      {
        type: "image",
        style: {
          type: image.extension,
          transformation: {
            width: image.width,
            height: image.height,
            ...style,
          },
        },
        data: image.buffer,
      },
    ];
  }

  computeStyle(image) {
    const style = {};

    if (style.width && !style.height) {
      const [numericWidth] = splitMeasure(width);
      style.width = numericWidth;
      style.height = numericWidth / image.ratio;
    } else if (!style.width && style.height) {
      const [numericHeight] = splitMeasure(height);
      style.height = numericHeight;
      style.width = numericHeight * image.ratio;
    }

    return style;
  }

  async getImage() {
    const src = this.node.attributes.src;
    console.log({ src })
    if (!src) return null;

    return await imageBase64ToBuffer(src);
  }

  supports() {
    return ["img"];
  }
}

export default ImgTransformer;
