import React from "react";

/**
 * TitanRoofLogo — SVG rendition of the TitanRoof mark. Two stacked "T"
 * glyphs with angular roof-peak tops form a monogram that reads as a
 * roofline: a solid block on the left and a hollow outline on the right.
 *
 * Both colors default to currentColor so a parent can render the mark
 * in white on a blue background (login / brand chip) or in slate on a
 * light surface (default) by simply setting `color`. Explicit `fill`
 * and `stroke` overrides remain available for fine control.
 */

interface Props {
  size?: number;
  className?: string;
  title?: string;
  fill?: string;
  stroke?: string;
}

const TitanRoofLogo: React.FC<Props> = ({
  size = 28,
  className,
  title = "TitanRoof",
  fill,
  stroke,
}) => {
  const aspect = 80 / 100;
  const fillColor = fill ?? "#1E2F3D";
  const strokeColor = stroke ?? fill ?? "#1E2F3D";

  return (
    <svg
      className={className}
      width={size}
      height={Math.round(size * aspect)}
      viewBox="0 0 100 80"
      fill="none"
      role="img"
      aria-label={title}
    >
      <title>{title}</title>
      {/* Left "T": filled block with sharp angular roof peak. Apex
          sits at ~46/8; roof slopes down to the right, stem is a
          vertical shaft that narrows slightly toward the bottom. */}
      <path
        d="M2 22 L46 4 L52 12 L46 20 L46 78 L24 78 L24 26 L2 26 Z"
        fill={fillColor}
      />
      {/* Right "T": hollow outline sharing the same angular peak
          silhouette but smaller and shifted right, forming the
          second letter of the monogram. */}
      <path
        d="M52 24 L88 8 L98 16 L98 22 L78 30 L78 78 L64 78 L64 34 L52 34 Z"
        fill="none"
        stroke={strokeColor}
        strokeWidth="3"
        strokeLinejoin="round"
        strokeLinecap="round"
      />
    </svg>
  );
};

export default TitanRoofLogo;
