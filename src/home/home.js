import React, { Component } from 'react';
import hello from 'hellojs/dist/hello.all.js';
import axios from 'axios';

class Home extends Component {
  constructor(props) {
    super(props);

    this.state = {
      contacts: []
    };

    this.onLogout = this.onLogout.bind(this);
  }

  componentDidMount() {
    const msft = hello('msft').getAuthResponse();

    axios.get(
      'https://graph.microsoft.com/v1.0/me/contacts?$select=displayName,emailAddresses',
      { headers: { Authorization: `Bearer ${msft.access_token}` }}
    ).then(res => {
      const contacts = res.data.value;
      console.log('res', res);
      this.setState({ contacts });
    });
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

        <button onClick={this.onLogout}>Logout</button>
      </div>
    );
  }
}

export default Home;
