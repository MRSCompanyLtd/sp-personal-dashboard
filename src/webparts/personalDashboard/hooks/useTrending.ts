import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IAadTokenProvider } from "../interfaces/IAadTokenProvider";
import { IUserTrending } from "../interfaces/IUserTrending";

interface IUseTrendingProps {
    context: WebPartContext;
}

interface IUseTrendingReturn {
    getTrending: () => Promise<void>;
    trending: IUserTrending[];
    loading: boolean;
}

const useTrending: (props: IUseTrendingProps) => IUseTrendingReturn = ({ context }) => {
    const [trending, setTrending] = React.useState<IUserTrending[]>([]);
    const [loading, setLoading] = React.useState<boolean>(false);

    const getTrending: () => Promise<void> = async () => {
        try {
            setLoading(true);
            const factory: IAadTokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
            const token: string = await factory.getToken('https://graph.microsoft.com');

            const trending: { id: string, value: IUserTrending[] } = await fetch('https://graph.microsoft.com/v1.0/me/insights/trending', {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${token}`
                }
            }).then((res: Response) => res.json());

            setTrending(trending.value);
            setLoading(false);
        }
        catch(e: unknown) {
            console.log(e);
            setLoading(false);
        }
    }

    return { getTrending, trending, loading }
}

export default useTrending;
