import * as React from 'react';
import { TextField, ITextFieldStyles, ITextFieldStyleProps } from 'office-ui-fabric-react/lib/TextField';
import { ILabelStyles, ILabelStyleProps } from 'office-ui-fabric-react/lib/Label';

interface IListTitleTextfieldProps {
  changeListTitle(title: string): void;
}

const ListTitleTextfield: React.FC<IListTitleTextfieldProps> = (props: IListTitleTextfieldProps) => {

  const handleChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    props.changeListTitle(event.target.value);
  };

  return <div>
    <TextField label="Your Sharepoint list title:" placeholder="list title"
      required styles={getStyles} onChange={handleChange} />
  </div>;
};

function getStyles(props: ITextFieldStyleProps): Partial<ITextFieldStyles> {
  return {
    subComponentStyles: {
      label: getLabelStyles,
    },
  };
}

function getLabelStyles(props: ILabelStyleProps): ILabelStyles {
  const { required } = props;
  return {
    root: required && {
      color: props.theme.palette.white
    },
  };
}

export default ListTitleTextfield;