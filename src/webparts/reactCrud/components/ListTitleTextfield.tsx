import * as React from 'react';
import { TextField, ITextFieldStyles, ITextFieldStyleProps } from 'office-ui-fabric-react/lib/TextField';
import { ILabelStyles, ILabelStyleProps } from 'office-ui-fabric-react/lib/Label';

interface IListTitleTextfieldProps {
  handleSubmit(title: string): void;
}

const ListTitleTextfield: React.FC<IListTitleTextfieldProps> = (props) => {

  const handleChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    props.handleSubmit(event.target.value);
  };

  return <div>
    <TextField label="Your Sharepoint list title:" placeholder="testList"
      required styles={getStyles} onChange={handleChange} />
  </div>;
};

function getStyles(props: ITextFieldStyleProps): Partial<ITextFieldStyles> {
  const { required } = props;
  return {
    // fieldGroup: [
    //    { width: 300 },
    //   required && {
    //     borderTopColor: props.theme.semanticColors.errorText,
    //   },
    // ],
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