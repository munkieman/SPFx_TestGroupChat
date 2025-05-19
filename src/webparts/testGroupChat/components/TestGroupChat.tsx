import * as React from 'react';
import type { ITestGroupChatProps } from './ITestGroupChatProps';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

// Chat component using React hooks (can be moved to its own file if desired)
const Chat: React.FC<ITestGroupChatProps> = (props) => {
  const { context } = props;

  const [userToAdd, setUserToAdd] = React.useState('');
  const [userToRemove, setUserToRemove] = React.useState('');
  const [members, setMembers] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [chatStatus, setChatStatus] = React.useState<string | null>(null);
  const [chatId, setChatId] = React.useState<string>(''); // <-- ChatId as state

  const getGraphClient = React.useCallback(async (): Promise<Client> => {
    const tokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
    return Client.init({
      authProvider: async (done) => {
        try {
          const token = await tokenProvider.getToken("https://graph.microsoft.com");
          done(null, token);
        } catch (err) {
          done(err, null);
        }
      }
    });
  }, [context.aadTokenProviderFactory]);

  const postMessageToChat = async (graphClient: any, chatId: string, message: string): Promise<void> => {
    try {
      const messagePayload = {
        body: {
          content: message
        }
      };

      await graphClient.api(`/chats/${chatId}/messages`).post(messagePayload);
      console.log('Message posted successfully');
    } catch (error) {
      console.error('Error posting message:', error);
    }
  };

  const createGroupChat = async (ownerUserId: string, chosenUserId: string): Promise<void> => {
    try {
      const client = await getGraphClient();

      const chatPayload ={
        chatType: 'Group',
        topic: "Test Chat",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${ownerUserId}`,
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${chosenUserId}`,
          }
        ],
        visibleHistoryStartDateTime: new Date().toISOString()
      };

      const response = await client.api(`/chats`).post(chatPayload);      
      console.log('Chat created successfully:', response);

      setChatId(response.id); // Set chatId in state

      await postMessageToChat(client, response.id, "Welcome to the chat! Let’s get started.");
       
      setChatStatus('Chat created successfully!');
    } catch (error) {
      console.error('Error creating chat:', error);
      setChatStatus('Failed to create chat.');
    }
  };

  const handleStartChat = () => {
    const ownerUserId = '63ba8e24-e214-4825-94f2-219a24addd23';
    const chosenUserId = '44929a9b-34a1-4091-9111-fa6e06b51665';
    createGroupChat(ownerUserId, chosenUserId);
  };

  const refreshMembers = React.useCallback(async () => {
    if (!chatId) {
      setMembers([]);
      return;
    }
    setLoading(true);
    try {
      const graphClient = await getGraphClient();
      const result = await graphClient.api(`/chats/${chatId}/members`).get();
      setMembers(result.value || []);
    } catch (e) {
      setMembers([]);
    }
    setLoading(false);
  }, [getGraphClient, chatId]);

  React.useEffect(() => {
    if (chatId) {
      refreshMembers();
    }
  }, [refreshMembers, chatId]);

  const addUser = async () => {
    if (!userToAdd || !chatId) return;
    setLoading(true);
    try {
      const graphClient = await getGraphClient();
      const now = new Date().toISOString();
      const memberPayload = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": ["Owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${userToAdd}`,
        "visibleHistoryStartDateTime": now
      };
      await graphClient.api(`/chats/${chatId}/members`).post(memberPayload);
      setUserToAdd('');
      await refreshMembers();
      alert('User added without history!');
    } catch (error) {
      alert('Error adding user');
      // eslint-disable-next-line no-console
      console.error(error);
    }
    setLoading(false);
  };

  const removeUser = async () => {
    if (!userToRemove || !chatId) return;
    setLoading(true);
    try {
      const graphClient = await getGraphClient();
      const membersResult = await graphClient.api(`/chats/${chatId}/members`).get();
      const memberToRemove = membersResult.value.find((m: any) => m.userId === userToRemove);
      if (memberToRemove) {
        await graphClient.api(`/chats/${chatId}/members/${memberToRemove.id}`).delete();
        setUserToRemove('');
        await refreshMembers();
        alert('User removed!');
      } else {
        alert('User not found in chat');
      }
    } catch (error) {
      alert('Error removing user');
      // eslint-disable-next-line no-console
      console.error(error);
    }
    setLoading(false);
  };

  return (
    <div>
      <div>Chat ID : {chatId}</div>
      <div>
        <PrimaryButton 
          text="Start Group Chat"
          onClick={handleStartChat}
        />
        {chatStatus && <p>{chatStatus}</p>}
      </div>

      <div style={{ marginBottom: 12 }}>
        <input
          type="text"
          placeholder="Azure AD object ID to add"
          value={userToAdd}
          onChange={e => setUserToAdd(e.target.value)}
          disabled={loading}
        />
        <PrimaryButton
          text="Add User Without History"
          onClick={addUser}
          disabled={loading || !userToAdd || !chatId}
          style={{ marginLeft: 8 }}
        />
      </div>
      <div style={{ marginBottom: 12 }}>
        <input
          type="text"
          placeholder="Azure AD object ID to remove"
          value={userToRemove}
          onChange={e => setUserToRemove(e.target.value)}
          disabled={loading}
        />
        <DefaultButton
          text="Remove User"
          onClick={removeUser}
          disabled={loading || !userToRemove || !chatId}
          style={{ marginLeft: 8 }}
        />
      </div>
      <div>
        <h4>Current Group Chat Members</h4>
        {loading ? <div>Loading members...</div> : null}
        <ul>
          {members.map(m => (
            <li key={m.id}>
              {m.displayName || m.userId} ({m.roles?.join(', ') || 'Member'})
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
};

export default Chat;