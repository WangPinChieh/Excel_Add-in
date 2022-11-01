import { Stack } from "@fluentui/react";
import * as React from "react";

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

const Header = (props: HeaderProps) => {
  const { title, logo, message } = props;

  return (
    <Stack horizontal>
      <img width="90" height="90" src={logo} alt={title} title={title} />
      <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{message}</h1>
    </Stack>
  );
};

export default Header;
