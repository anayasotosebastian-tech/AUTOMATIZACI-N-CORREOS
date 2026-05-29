{ pkgs }: {
  deps = [
    pkgs.python311
    pkgs.python311Packages.pip
    pkgs.libreoffice-unwrapped
    pkgs.fontconfig
    pkgs.freetype
  ];
}
