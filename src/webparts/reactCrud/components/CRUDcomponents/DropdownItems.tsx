import * as React from 'react';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

const dropdownStyles: Partial<IDropdownStyles> = {
  label: { color: "white" },
};

interface IDropdownItemsProps {
  items: IDropdownOption[];
  chooseItem(item: any): void;
}


const stackTokens: IStackTokens = { childrenGap: 20 };

const DropdownItems: React.FC<IDropdownItemsProps> = (props: IDropdownItemsProps) => {

  const handleChange = (event: React.ChangeEvent<HTMLInputElement>, item: any) => {
    props.chooseItem(item);
  };

  return (
    <Stack tokens={stackTokens}>
      <Dropdown
        placeholder="list items"
        label="Your Sharepoint list items:"
        options={props.items}
        defaultSelectedKey={0}
        styles={dropdownStyles}
        onChange={handleChange}
      />
    </Stack>
  );

};

export default DropdownItems;