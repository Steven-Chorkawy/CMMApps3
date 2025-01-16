import { Shimmer } from '@fluentui/react';
import * as React from 'react';


export class MyShimmer extends React.Component {
    public render(): React.ReactElement<any> {
        const style = { margin: '10px' };
        return <div>
            <Shimmer style={style} />
            <Shimmer width="75%" style={style} />
            <Shimmer width="50%" style={style} />
        </div>;
    }
}
