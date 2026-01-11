import { AccountInfo } from '@azure/msal-node';

export interface TokenData {
  accessToken: string;
  refreshToken: string; // MSALでは使用しないが互換性のため保持
  expiresOn: Date;
  account: AccountInfo;
}

export interface UserTokenData {
  userId: string;
  accessToken: string;
  refreshToken: string; // MSALでは使用しない
  expiresOn: Date;
  account: AccountInfo;
  subscriptions: SubscriptionData[];
}

export interface SubscriptionData {
  id: string;
  expirationDateTime: string;
  resource: string;
}

export interface Subscription {
  id: string;
  resource: string;
  changeType: string;
  notificationUrl: string;
  expirationDateTime: string;
  clientState: string;
}

export interface GraphNotification {
  subscriptionId: string;
  subscriptionExpirationDateTime: string;
  changeType: string;
  resource: string;
  resourceData: {
    '@odata.type': string;
    '@odata.id': string;
    '@odata.etag': string;
    id: string;
  };
  clientState: string;
  tenantId: string;
}

export interface NotificationPayload {
  value: GraphNotification[];
}

export interface EmailMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  body: {
    contentType: string;
    content: string;
  };
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  receivedDateTime: string;
  hasAttachments: boolean;
  webLink?: string;
}

export interface StoredEmailMessage {
  id: string;
  userId: string;
  subscriptionId: string;
  subject: string;
  from: string;
  fromAddress: string;
  bodyPreview: string;
  bodyContent: string;
  receivedDateTime: string;
  notificationReceived: string;
  hasAttachments: boolean;
  webLink?: string;
}