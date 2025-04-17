export interface ITileProps {
  header: string;
  text: string;
}

export interface ISpFourTilesProps {
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
  tile1: ITileProps;
  tile2: ITileProps;
  tile3: ITileProps;
  tile4: ITileProps;
  displayMode: number;
  updateProperty: (propertyPath: string, newValue: string) => void;
}
