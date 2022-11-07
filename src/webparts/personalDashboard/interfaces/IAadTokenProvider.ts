export interface IAadTokenProvider {
    getToken: (resourceEndpoint: string, useCachedToken?: boolean) => Promise<string>
}