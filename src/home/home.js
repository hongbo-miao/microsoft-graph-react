import React, { Component } from 'react';
import hello from 'hellojs/dist/hello.all.js';
import axios from 'axios';

class Home extends Component {
  constructor(props) {
    super(props);

    const msft = hello('msft').getAuthResponse();

    this.state = {
      contacts: [],
      token: msft.access_token
    };

    this.onWriteToExcel = this.onWriteToExcel.bind(this);
    this.onLogout = this.onLogout.bind(this);
  }

  componentDidMount() {
    const { token } = this.state;

    axios.get(
      'https://graph.microsoft.com/v1.0/me/contacts?$select=displayName,emailAddresses',
      { headers: { Authorization: `Bearer ${token}` }}
    ).then(res => {
      const contacts = res.data.value;
      this.setState({ contacts });
    });
  }

  onWriteToExcel() {
    const { token, contacts } = this.state;

    const values = [];

    contacts.forEach(contact => {
      values.push([contact.displayName, contact.emailAddresses[0].address]);
    });

    axios
      .post('https://graph.microsoft.com/v1.0/me/drive/root:/demo.xlsx:/workbook/tables/Table1/rows/add',
        { index: null, values },
        { headers: { Authorization: `Bearer ${token}` }}
      )
      .then(res => console.log(res))
      .catch(err => console.error(err));
  }

  onLogout() {
    hello('msft').logout().then(
      () => this.props.history.push('/'),
      e => console.error(e.error.message)
    );
  }

  renderContacts() {
    const { contacts } = this.state;

    return (
      contacts.map(contact =>
        <tr key={contact.id}>
          <td>{contact.displayName}</td>
          <td>{contact.emailAddresses[0].address}</td>
        </tr>
      )
    );
  }

  render() {
    return (
      <div>
        <table>
          <thead>
            <tr>
              <th>Name</th>
              <th>Email</th>
            </tr>
          </thead>
          <tbody>
            {this.renderContacts()}
          </tbody>
        </table>

        <button onClick={this.onWriteToExcel}>Write to Excel</button>
        <button onClick={this.onLogout}>Logout</button>
      </div>
    );
  }
}

export default Home;
