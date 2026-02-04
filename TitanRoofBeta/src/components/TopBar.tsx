import React from "react";

interface TopBarProps {
  label: string;
}

const TopBar: React.FC<TopBarProps> = ({ label }) => {
  return <div className="topBar">{label}</div>;
};

export default TopBar;
