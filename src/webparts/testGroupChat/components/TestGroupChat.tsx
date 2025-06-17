import * as React from 'react';
import { useState, useEffect } from 'react';
import type { ITestGroupChatProps } from './ITestGroupChatProps';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

// Chat component using React hooks (can be moved to its own file if desired)
const Chat: React.FC<ITestGroupChatProps> = (props) => {
  const { 
    userDisplayName,
    context,
    owners=[],
    //currentUserEmail
  } = props;

  //const [members, setMembers] = useState<any[]>([]);
  const [displayName, setDisplayName] = useState(userDisplayName);
  //const [presences, setPresences] = useState<{ [objectId: string]: { name: string, presence: string } }>({});
  const [loading, setLoading] = useState(false);
  const [chatStatus, setChatStatus] = useState<string | null>(null);
  const [chatId, setChatId] = useState<string>(''); // <-- ChatId as state
  const [exporting, setExporting] = useState(false);

/*  
  const getPresenceStyle = (presence: string): React.CSSProperties => {
    const normalized = (presence || '').toLowerCase();
    if (normalized === 'available') {
      return { color: 'green', fontWeight: 'bold' };
    }
    if (normalized === 'busy' || normalized === 'focusing' || normalized === 'donotdisturb') {
      return { color: 'red', fontWeight: 'bold' };
    }
    if (normalized === 'unavailable') {
      return { color: 'grey' };
    }
    return {};
  };
*/

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

  const createGroupChat = async (): Promise<void> => {

    try {
      const graphClient = await getGraphClient();
      const aadClient: AadHttpClient = await context.aadHttpClientFactory.getClient('https://graph.microsoft.com');

      /* Munkie 365 
      const userIds = [
        'c79cdecd-0a47-483a-a55b-e5612be126f0',
        'f6e0e5fd-46a5-4c6e-b42b-13ec7fdc8c0f'
        // add more
      ]; */

      /* Max Dev 365 */
      //const userIds = [
      //  '8532aff4-a77d-4bde-9657-36cd12269f38',
      //  '2939cad2-59eb-4e66-82da-6f9e47f1e142'
      //]

      const objectIds: string[] = [];
      await Promise.all(
        (owners || []).map(async (owner) => {
          const upn = owner.email || (owner.loginName && owner.loginName.split('|').pop());
          if (!upn) return;
          try {
            const userResponse: HttpClientResponse = await aadClient.get(
              `https://graph.microsoft.com/v1.0/users/${upn}`,
              AadHttpClient.configurations.v1
            );
            if (!userResponse.ok) throw new Error("Failed to fetch user");
            const userData = await userResponse.json();
            if (userData.id) objectIds.push(userData.id);
          } catch {
            // skip if cannot resolve
          }
        })
      );

      if (objectIds.length === 0) {
        setChatStatus('No users selected.');
        return;
      }

      const currentUser = await graphClient.api('/me').get();
      const currentUserId = currentUser.id;

      // Ensure current user is in the userIds array
      if (!objectIds.includes(currentUserId)) {
        objectIds.push(currentUserId);
      }

      // Prepare members array for the chat payload
      const members = objectIds.map(uid => ({
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${uid}`
      }));

      const chatPayload = {
        chatType: 'Group',
        topic: "Expenses Chat with "+displayName,
        members,
        visibleHistoryStartDateTime: new Date().toISOString()
      };

      const response = await graphClient.api(`/chats`).post(chatPayload);      
      console.log('Chat created successfully:', response);

      setChatId(response.id); // Set chatId in state

      await postMessageToChat(graphClient, response.id, "Welcome to the chat! Let’s get started.");
       
      setChatStatus('Chat created successfully!');
    } catch (error) {
      console.error('Error creating chat:', error);
      setChatStatus('Failed to create chat.');
    }
  };

  const refreshMembers = React.useCallback(async () => {
    setLoading(true);
    try {
      // Use the owners array from props directly
      //setMembers(owners || []);
    } catch (e) {
      //setMembers([]);
    }
    setLoading(false);
  }, [owners]);

  // Remove all members from the group chat
  const removeAllMembers = async () => {
    if (!chatId) return;
    setLoading(true);
    try {
      const graphClient = await getGraphClient();
      const membersResult = await graphClient.api(`/chats/${chatId}/members`).get();
      const membersList = membersResult.value;
      for (const member of membersList) {
        await graphClient.api(`/chats/${chatId}/members/${member.id}`).delete();
      }
      //setMembers([]);
      alert('All members removed!');
    } catch (error) {
      alert('Error removing members');
      console.error(error);
    }
    setLoading(false);
  };

/*  
  const fetchPresences = async () => {
    if (!owners || owners.length === 0) {
      setPresences({});
      return;
    }
    
    const aadClient: AadHttpClient = await context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
    const presenceResult: { [objectId: string]: { name: string, presence: string } } = {};

    await Promise.all(
      owners.map(async (owner) => {
        
        // Get UPN/email from owner
        const upn = owner.email || (owner.loginName && owner.loginName.split('|').pop());
        console.log("Processing owner:", owner, "UPN:", upn);
        if (!upn) return;
        
        try {
            // 1. Get the user object from Graph to obtain the object ID
            const userResponse: HttpClientResponse = await aadClient.get(
              `https://graph.microsoft.com/v1.0/users/${upn}`,
              AadHttpClient.configurations.v1
            );
            if (!userResponse.ok) throw new Error("Failed to fetch user");
            const userData = await userResponse.json();
            const objectId = userData.id;
            const displayName = userData.displayName || owner.text;

            console.log("Owners:", owners);
            console.log("Fetching presence for:", userData.displayName, "Object ID:", objectId);
                       
            // 2. Get the presence using the object ID
            let presenceValue = "Unknown";
            try {
              const presenceResponse: HttpClientResponse = await aadClient.get(
                `https://graph.microsoft.com/v1.0/users/${objectId}/presence`,
                AadHttpClient.configurations.v1
              );
              if (presenceResponse.ok) {
                const presenceData = await presenceResponse.json();
                presenceValue = presenceData.availability || "Unknown";
                console.log("Presence response:", presenceData);
                console.log("Final presences state:", presenceResult);
              }
            } catch {
              // fall through, leave as "Unknown"
            }
            presenceResult[objectId] = { name: displayName, presence: presenceValue };
          } catch {
            // If the user can't be resolved, don't add it
          }
      })
    );
    setPresences(presenceResult);
  };
*/

  // Export chat conversation as text file
  const exportChat = async () => {
    if (!chatId) return;
    setExporting(true);
    try {
      const graphClient = await getGraphClient();
      const messages: any[] = [];
      let response = await graphClient.api(`/chats/${chatId}/messages`).get();

      messages.push(...response.value);

      // Handle paging if more messages
      while (response["@odata.nextLink"]) {
        response = await graphClient.api(response["@odata.nextLink"]).get();
        messages.push(...response.value);
      }

      // Format messages
      const text = messages.map(msg => {
        const time = new Date(msg.createdDateTime).toLocaleString();
        const user = msg.from?.user?.displayName || "Unknown";
        // Remove HTML tags from content
        const content = (msg.body?.content || "").replace(/<[^>]+>/g, '');
        return `[${time}] ${user}: ${content}`;
      }).join('\n');

      // Download as .txt
      const blob = new Blob([text], { type: 'text/plain' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = `chat-${chatId}.txt`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      alert('Error exporting chat');
      console.error(error);
    }
    setExporting(false);
  };

  // Update displayName if userDisplayName prop changes
  useEffect(() => {
    setDisplayName(userDisplayName);
  }, [userDisplayName]);

  // Fetch user presence when owners change
  //useEffect(() => {
  //  if (owners && owners.length > 0) {
  //    fetchPresences();
  //  } else {
  //    setPresences({});
  //  }
  //}, [owners]);

  //useEffect(() => {
  //  console.log("owners prop changed:", owners);
  //  fetchPresences();   
  //}, [owners, context.aadHttpClientFactory]);

  useEffect(() => {
    refreshMembers();
  }, [refreshMembers]);

  return (
    <div>
      <div style={{ marginBottom: 10 }}>
        <PrimaryButton 
          text="Start Group Chat"
          onClick={createGroupChat}
          style={{backgroundColor: '#502e91', border: 'none' }}
        />
      </div>

      <div style={{ marginBottom: 10 }}>
        <PrimaryButton
          text="Remove All Members"
          onClick={removeAllMembers}
          disabled={loading || !chatId}
          style={{ backgroundColor: '#f44336', border: 'none' }}
        />
      </div>

      <div style={{ marginBottom: 10 }}>
        <PrimaryButton
          text={exporting ? "Exporting..." : "Export Chat as Text"}
          onClick={exportChat}
          disabled={exporting || !chatId}
          style={{backgroundColor: '#200941', border: 'none' }}
        />
      </div>

      <div>Chat ID : {chatId}</div>
      {chatStatus && <p>{chatStatus}</p>}
      <br/>

      <div>
        <h4>Current Group Chat Members</h4>
        {loading ? <div>Loading members...</div> : null}
          <h3>Advisors</h3>
          
      </div>
    </div>
  );
};

export default Chat;

/*

          <ul>
            {members.length > 0 ? (
              members.map((member, idx) => {
                // Try to get the objectId for presence lookup
                const objectId = member.id || member.objectId;
                const info = objectId ? presences[objectId] : undefined;
                const name = member.displayName || member.text || member.email || 'Unknown';
                const presence = info ? info.presence : 'Unknown';
                return (
                  <li key={objectId || member.email || idx}>
                    {name} — <span style={getPresenceStyle(presence)}>{presence}</span>
                  </li>
                );
              })
            ) : (
              <li>No advisors selected.</li>
            )}
          </ul>
          <ul>
            {Object.keys(presences).length > 0 ? (
              Object.keys(presences).map(objectId => {
                const info = presences[objectId];
                return (
                  <li key={objectId}>
                    {info.name} — <span style={getPresenceStyle(info.presence)}>{info.presence}</span>
                  </li>
                );
              })
            ) : (
              <li>No advisors selected.</li>
            )}
          </ul> 


  const addUser = async () => {
    if (!userToAdd || !chatId) return;
    setLoading(true);
    try {
      const graphClient = await getGraphClient();
      const now = new Date().toISOString();
      const userEmail = props.context.pageContext.user.email;
        
      //Fetch the user ID
      const userResponse = await graphClient.api(`/users/${userEmail}`).get();  
      const userData = await userResponse.json();
      const userId = userData.id;
      setUserToAdd(userId);
      console.log("userID",userId,userData);

      const memberPayload = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": ["Owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${userToAdd}`,
        "visibleHistoryStartDateTime": now
      };

      await graphClient.api(`/chats/${chatId}/members`).post(memberPayload);      
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

*/