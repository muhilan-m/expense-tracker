import * as React from 'react';
import { TextField, PrimaryButton } from 'office-ui-fabric-react';
//import { sp } from "@pnp/sp/presets/all";  



export interface ExpenseFormProps {
  onSubmit: (formData: ExpenseFormData) => void;
}

export interface ExpenseFormData {
  name: string;
  department: string;
  projectName: string;
  expense: string;
  remarks: string;
}

export interface ExpenseFormState {
  formData: ExpenseFormData;
}

export default class ExpenseForm extends React.Component<ExpenseFormProps, ExpenseFormState> {
  constructor(props: ExpenseFormProps) {
    super(props);

    this.state = {
      formData: {
        name: '',
        department: '',
        projectName: '',
        expense: '',
        remarks: ''
      }
    };
  }

  handleInputChange = (fieldName: keyof ExpenseFormData, value: string) => {
    const { formData } = this.state;
    this.setState({
      formData: {
        ...formData,
        [fieldName]: value
      }
    });
  };

    

  handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const { formData } = this.state;
  
    try {
      // Get the SharePoint context and initialize the Web object
      const context = sp.web.getContext();
      const web = new Web(context.pageContext.web.absoluteUrl);
  
      // Create a new item in the SharePoint list
      await web.lists.getByTitle('YourListTitle').items.add(formData);
  
      // Invoke the onSubmit callback with the form data
      this.props.onSubmit(formData);
    } catch (error) {
      // Handle any errors that occur during the submission
      console.log('Error creating record:', error);
    }
  };
  

  render() {
    const { formData } = this.state;

    return (
      <form onSubmit={this.handleSubmit}>
        <TextField
          label="Name"
          value={formData.name}
          required
          onChanged={(value) => this.handleInputChange('name', value)}
        />

        <TextField
          label="Department"
          value={formData.department}
          required
          onChanged={(value: string) => this.handleInputChange('department', value)}
        />

        <TextField
          label="Project Name"
          value={formData.projectName}
          required
          onChanged={(value: string) => this.handleInputChange('projectName', value)}
        />

        <TextField
          label="Expense"
          value={formData.expense}
          required
          onChanged={(value: string) => this.handleInputChange('expense', value)}
        />

        <TextField
          label="Remarks"
          multiline
          value={formData.remarks}
          required
          onChanged={(value: string) => this.handleInputChange('remarks', value)}
        />

        <PrimaryButton type="submit" text="Submit" />
      </form>
    );
  }
}
