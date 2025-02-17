import * as React from 'react';
import { IMergeandimageProps } from './IMergeandimageProps';

const Mergeandimage: React.FunctionComponent<IMergeandimageProps> = (props) => {
  const { description, graphClient } = props;

  return (
    <div>
      <h1>{description}</h1>
      {/* Example usage of graphClient */}
      <button onClick={() => graphClient.api('/me').get()}>
        Fetch User Info
      </button>
    </div>
  );
};

export default Mergeandimage;
