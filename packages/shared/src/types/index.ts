// Shared type definitions for frontend and backend

/**
 * Graph API subscription related types
 */
export interface Subscription {
  id?: string;
  resource: string;
  changeType: string;
  notificationUrl: string;
  expirationDateTime: string;
  clientState?: string;
}

/**
 * Change notification from Graph API
 */
export interface ChangeNotification {
  subscriptionId: string;
  subscriptionExpirationDateTime: string;
  changeType: string;
  resource: string;
  resourceData?: {
    '@odata.type': string;
    '@odata.id': string;
    id: string;
  };
  tenantId?: string;
}

/**
 * User token information for OBO flow
 */
export interface UserTokenInfo {
  accessToken: string;
  userId: string;
  tenantId: string;
}

/**
 * API response wrapper
 */
export interface ApiResponse<T = unknown> {
  success: boolean;
  data?: T;
  error?: string;
  message?: string;
}