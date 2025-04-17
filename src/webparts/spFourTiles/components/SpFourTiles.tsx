import * as React from 'react';
import styles from './SpFourTiles.module.scss';
import type { ISpFourTilesProps, ITileProps } from './ISpFourTilesProps';
import { DisplayMode } from '@microsoft/sp-core-library';

interface ITileComponentProps {
  tile: ITileProps;
  displayMode: number;
  onHeaderChange: (newValue: string) => void;
  onTextChange: (newValue: string) => void;
}

const Tile: React.FC<ITileComponentProps> = (props) => {
  const { tile, displayMode, onHeaderChange, onTextChange } = props;
  const isEdit = displayMode === DisplayMode.Edit;

  const handleHeaderChange = (event: React.FormEvent<HTMLDivElement>) => {
    onHeaderChange(event.currentTarget.textContent || '');
  };

  const handleTextChange = (event: React.FormEvent<HTMLDivElement>) => {
    onTextChange(event.currentTarget.textContent || '');
  };

  return (
    <div className={styles.tile}>
      {isEdit ? (
        <>
          <div
            className={`${styles.editableField} ${styles.headerField}`}
            contentEditable={true}
            onBlur={handleHeaderChange}
            dangerouslySetInnerHTML={{ __html: tile.header }}
          />
          <div
            className={`${styles.editableField} ${styles.textField}`}
            contentEditable={true}
            onBlur={handleTextChange}
            dangerouslySetInnerHTML={{ __html: tile.text }}
          />
        </>
      ) : (
        <>
          <div className={styles.tileHeader}>{tile.header}</div>
          <div className={styles.tileText}>{tile.text}</div>
        </>
      )}
    </div>
  );
};

export default class SpFourTiles extends React.Component<ISpFourTilesProps> {
  public render(): React.ReactElement<ISpFourTilesProps> {
    const {
      hasTeamsContext,
      displayMode,
      tile1,
      tile2,
      tile3,
      tile4
    } = this.props;

    const updateTileProperty = (tileName: string, propertyName: string, newValue: string) => {
      this.props.updateProperty(`${tileName}.${propertyName}`, newValue);
    };

    return (
      <section className={`${styles.spFourTiles} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.tilesContainer}>
          <Tile 
            tile={tile1} 
            displayMode={displayMode} 
            onHeaderChange={(newValue) => updateTileProperty('tile1', 'header', newValue)} 
            onTextChange={(newValue) => updateTileProperty('tile1', 'text', newValue)} 
          />
          <Tile 
            tile={tile2} 
            displayMode={displayMode} 
            onHeaderChange={(newValue) => updateTileProperty('tile2', 'header', newValue)} 
            onTextChange={(newValue) => updateTileProperty('tile2', 'text', newValue)} 
          />
          <Tile 
            tile={tile3} 
            displayMode={displayMode} 
            onHeaderChange={(newValue) => updateTileProperty('tile3', 'header', newValue)} 
            onTextChange={(newValue) => updateTileProperty('tile3', 'text', newValue)} 
          />
          <Tile 
            tile={tile4} 
            displayMode={displayMode} 
            onHeaderChange={(newValue) => updateTileProperty('tile4', 'header', newValue)} 
            onTextChange={(newValue) => updateTileProperty('tile4', 'text', newValue)} 
          />
        </div>
      </section>
    );
  }
}
