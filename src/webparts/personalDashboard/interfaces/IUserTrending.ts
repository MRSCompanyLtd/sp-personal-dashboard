export interface IUserTrending {
    resourceReference: {
        id: string;
        type: string;
        webUrl: string;
    }
    resourceVisualization: {
        containerDisplayName: string;
        containerType: string;
        containerWebUrl: string;
        mediaType: string;
        previewImageUrl: string;
        previewText: string;
        title: string;
        type: string;
    }
}