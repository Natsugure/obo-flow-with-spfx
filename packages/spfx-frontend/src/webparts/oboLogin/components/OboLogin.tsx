import * as React from 'react';
import styles from './OboLogin.module.scss';
import type { IOboLoginProps } from './IOboLoginProps';
import type { IOboLoginState } from './IOboLoginState';

import { DefaultButton, PrimaryButton } from '@fluentui/react';
import { AadHttpClient, IHttpClientOptions } from '@microsoft/sp-http';

export default class OboLogin extends React.Component<IOboLoginProps, IOboLoginState> {
  constructor(props: IOboLoginProps) {
    super(props);

    this.state = {
      data: null,
      isLoading: false,
      error: null,
    };
  }

  public render(): React.ReactElement<IOboLoginProps> {
    const {
      isDarkTheme,
      hasTeamsContext,
      apiUrl,
      apiClient
    } = this.props;

    const {
      data,
      isLoading,
      error
    } = this.state;

    return (
      <section className={`${styles.oboLogin} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>メール通知システム 利用開始ポータル</h2>
        </div>
        <div className={styles.oboLogin}>
          <PrimaryButton text="ログイン" onClick={this._sendCredential} disabled={isLoading} />

          {/* エラー表示 */}
          {error && (
            <div className={styles.error}>
              <p>エラー: {error}</p>
            </div>
          )}

          {/* データ表示 */}
          {data && (
            <div className = {styles.success}>
              <h3>登録成功</h3>
              <p>ユーザーID: {data.userId}</p>
              <p>サブスクリプションID: {data.subscriptionId}</p>
              {data.expirationDateTime && (
                <p>有効期限: { new Date(data.expirationDateTime).toLocaleDateString('ja-jp')}</p>
              )}
            </div>
          )}
        </div>
      </section>
    );
  }

  private _sendCredential = async (): Promise<void> => {
    const {
      apiUrl,
      apiClient,
    } = this.props;

    this.setState({ isLoading: true, error: null });

    try {
      console.log('[OboLogin] Starting API call to:', `${apiUrl}/api/subscribe`);

      const response = await apiClient.post(`${apiUrl}/api/subscribe`, AadHttpClient.configurations.v1, {});

      console.log('[OboLogin] Response status:', response.status);
      console.log('[OboLogin] Response ok:', response.ok);

      if (!response.ok) {
        // レスポンスボディを読み取る
        let errorDetail = response.statusText;
        try {
          const errorBody = await response.json();
          errorDetail = errorBody.error || errorBody.message || response.statusText;
          console.error('[OboLogin] Error response body:', errorBody);
        } catch (e) {
          // JSONパースに失敗した場合はテキストとして読み取る
          try {
            errorDetail = await response.text();
            console.error('[OboLogin] Error response text:', errorDetail);
          } catch (e2) {
            console.error('[OboLogin] Could not read error response');
          }
        }
        throw new Error(`API call failed (${response.status}): ${errorDetail}`);
      }

      const data = await response.json();
      console.log('[OboLogin] Success! Data:', data);
      this.setState({ data, isLoading: false });

    } catch (error: any) {
      console.error('[OboLogin] Error caught:', error);
      this.setState({
        error: error.message || 'Something went wrong.',
        isLoading: false,
      });
    }
  }
}
