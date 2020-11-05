import * as React from 'react';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { ILabelStyles, ILabelStyleProps } from 'office-ui-fabric-react/lib/Label';

interface IListTitleTextfieldProps {
  changeItemTitle(title: string): void;
  title: string;
}

const ListTitleTextfield: React.FC<IListTitleTextfieldProps> = (props: IListTitleTextfieldProps) => {

  const handleChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    props.changeItemTitle(event.target.value);
  };

  return <div>
    <TextField label="Write your new:" placeholder="new-item"
      styles={getStyles} onChange={handleChange} value={props.title} />
  </div>;
};

function getStyles(): Partial<ITextFieldStyles> {
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