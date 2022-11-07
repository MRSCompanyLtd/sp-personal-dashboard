import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUserLink } from "../interfaces/IUserLink";
import { IAadTokenProvider } from "../interfaces/IAadTokenProvider";

interface IUseLinksProps {
    context: WebPartContext;
}

interface IUseLinksReturn {
    links: IUserLink[];
    getLinks: () => Promise<void>;
    updateLinks: (newLink: IUserLink) => Promise<void>;
    deleteLink: (link: IUserLink) => Promise<void>;
    loading: boolean;
}

const useLinks: (props: IUseLinksProps) => IUseLinksReturn = ({ context }) => {
    const [links, setLinks] = React.useState<IUserLink[]>([]);
    const [loading, setLoading] = React.useState<boolean>(false);

    const getLinks = async (): Promise<void> => {
        try {
            setLoading(true);
            const factory: IAadTokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
            const token: string = await factory.getToken('https://graph.microsoft.com');

            const apps: { Links: IUserLink[] } = await fetch('https://graph.microsoft.com/v1.0/me/extensions/MyLinks', {
                headers: {
                    'Authorization': `Bearer ${token}`
                }
            }).then((res: Response) => res.json());
            
            const parsed: IUserLink[] = apps.Links.map((res: IUserLink) => {
                return {
                    name: res.name,
                    description: res.description,
                    url: res.url
                }
            });

            setLinks(parsed);
            setLoading(false);
        }
        catch(e: unknown) {
            console.log(e);
        }
    }

    const updateLinks = async (newLink: IUserLink): Promise<void> => {
        try {
            const factory: IAadTokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
            const token: string = await factory.getToken('https://graph.microsoft.com');

            const output: boolean | Response = await fetch('https://graph.microsoft.com/v1.0/me/extensions/MyLinks', {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${token}`
                }
            }).catch(() => {
                return false;
            });

            if (!output) {
                await fetch('https://graph.microsoft.com/v1.0/me/extensions', {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${token}`
                    },
                    body: JSON.stringify({
                        extensionName: 'MyLinks',
                        Links: {
                            name: newLink.name,
                            url: newLink.url,
                            description: newLink.description                            
                        }
                    })
                });
            } else {
                await fetch('https://graph.microsoft.com/v1.0/me/extensions/MyLinks', {
                    method: 'PATCH',
                    headers: {
                        'Authorization': `Bearer ${token}`,
                    },
                    body: JSON.stringify({
                        extensionName: 'MyLinks',
                        Links: [
                            ...links,
                            newLink
                        ]
                    })
                });
            }

            await getLinks();
        }
        catch(e: unknown) {
            console.log(e);
        }
    }

    const deleteLink = async (link: IUserLink): Promise<void> => {
        try {
            setLoading(true);
            const factory: IAadTokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
            const token: string = await factory.getToken('https://graph.microsoft.com');

            const updatedLinks: IUserLink[] = [...links];
            const index: number = updatedLinks.findIndex((l: IUserLink) => l.name === link.name);
            updatedLinks.splice(index, 1);

            await fetch('https://graph.microsoft.com/v1.0/me/extensions/MyLinks', {
                method: 'PATCH',
                headers: {
                    'Authorization': `Bearer ${token}`
                },
                body: JSON.stringify({
                    extensionName: 'MyLinks',
                    Links: updatedLinks
                })
            });

            setLinks(updatedLinks);
            setLoading(false);
        }
        catch(e: unknown) {
            console.log(e);

            setLoading(false);
        }
    }

    return { getLinks, updateLinks, deleteLink, links, loading }
}

export default useLinks;
