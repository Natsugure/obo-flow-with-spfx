export interface ISubscriptionData {
  userId: string;
  subscriptionId: string;
  expirationDateTime: string;
  resource?: string;
}

export interface IOboLoginState {
  data: ISubscriptionData | null;
  isLoading: boolean;
  error: string | null;
}
