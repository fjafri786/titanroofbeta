import React from "react";

/**
 * TitanRoofLogo — SVG rendition of the TitanRoof mark. Two stylized
 * "T" shapes form a stacked roof-peak monogram: a solid dark slate
 * block on the left and a hollow outline on the right, sharing a
 * shallow peak that reads as a roofline.
 *
 * The viewBox is normalized to 64×56 so the component slots cleanly
 * into toolbar rows alongside 20-28px icons without distortion.
 */

interface Props {
  size?: number;
  className?: string;
  title?: string;
}

const TitanRoofLogo: React.FC<Props> = ({ size = 28, className, title = "TitanRoof" }) => {
  const fill = "#1E2F3D";
  const stroke = "#4F6C7F";
  const aspect = 56 / 64;

  return (
    <svg
      className={className}
      width={size}
      height={Math.round(size * aspect)}
      viewBox="0 0 64 56"
      fill="none"
      role="img"
      aria-label={title}
    >
      <title>{title}</title>
      {/* Left "T": filled dark slate, angled roof peak */}
      <path
        d="M3 14 L26 2 L30 10 L26 16 L26 54 L14 54 L14 18 L3 18 Z"
        fill={fill}
      />
      {/* Right "T": hollow outline, lighter slate */}
      <path
        d="M30 18 L52 8 L61 16 L61 20 L49 24 L49 54 L39 54 L39 27 L30 27 Z"
        fill="none"
        stroke={stroke}
        strokeWidth="2.2"
        strokeLinejoin="round"
        strokeLinecap="round"
      />
    </svg>
  );
};

export default TitanRoofLogo;
