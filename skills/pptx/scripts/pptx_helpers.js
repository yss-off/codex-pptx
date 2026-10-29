'use strict';

// Coordinates are inches; imageWidth/imageHeight are positive source dimensions
// in the same unit (usually pixels). No filesystem, network, or global deck state.
function finite(value, name) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new TypeError(`${name} must be a finite number`);
  }
}

function dimensions(imageWidth, imageHeight, box) {
  finite(imageWidth, 'imageWidth');
  finite(imageHeight, 'imageHeight');
  if (imageWidth <= 0 || imageHeight <= 0) {
    throw new RangeError('Source image dimensions must be positive');
  }
  if (!box || typeof box !== 'object') throw new TypeError('box is required');
  for (const key of ['x', 'y', 'w', 'h']) finite(box[key], `box.${key}`);
  if (box.w <= 0 || box.h <= 0) throw new RangeError('Box dimensions must be positive');
}

/** Fit a complete image into the box, centered, without stretching or cropping. */
function imageContain(imageWidth, imageHeight, box) {
  dimensions(imageWidth, imageHeight, box);
  const scale = Math.min(box.w / imageWidth, box.h / imageHeight);
  const w = imageWidth * scale;
  const h = imageHeight * scale;
  return { x: box.x + (box.w - w) / 2, y: box.y + (box.h - h) / 2, w, h };
}

/** Fill the box by centered cropping, using PptxGenJS's supported cover sizing.
 * Keep the source aspect ratio in w/h: the exporter uses it to derive srcRect.
 * The final frame uses sizing.w/h. Choose this only when losing edges is intended.
 */
function imageCrop(imageWidth, imageHeight, box) {
  dimensions(imageWidth, imageHeight, box);
  // Cover needs only the source ratio. Normalize its longest edge to one inch:
  // PptxGenJS 4.0.1 treats numeric dimensions >=100 as EMU, so virtual cover
  // extents (e.g. 120x4 for a 30:1 image) would mix units and reverse the crop.
  const longestEdge = Math.max(imageWidth, imageHeight);
  return {
    x: box.x, y: box.y, w: imageWidth / longestEdge, h: imageHeight / longestEdge,
    sizing: { type: 'cover', w: box.w, h: box.h },
  };
}

/** Directed line from start to end. endArrowType stays on the requested endpoint.
 * Geometry wins over accidental x/y/w/h/flip values in options. Rotation is
 * rejected because rotating afterwards would no longer connect the endpoints.
 */
function lineBetween(x1, y1, x2, y2, options = {}) {
  for (const [name, value] of Object.entries({ x1, y1, x2, y2 })) finite(value, name);
  if (x1 === x2 && y1 === y2) throw new RangeError('A line needs distinct endpoints');
  if (!options || typeof options !== 'object' || Array.isArray(options)) {
    throw new TypeError('options must be an object');
  }
  if (options.rotate !== undefined && options.rotate !== 0) {
    throw new RangeError('Do not rotate an endpoint-defined line');
  }
  return {
    ...options,
    x: Math.min(x1, x2), y: Math.min(y1, y2),
    w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1,
  };
}

module.exports = { imageContain, imageCrop, lineBetween };
