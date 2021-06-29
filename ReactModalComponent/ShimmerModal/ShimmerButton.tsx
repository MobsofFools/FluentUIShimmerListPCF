import * as React from "react";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { Modal } from "@fluentui/react/lib/Modal";
import { IInputs } from "../generated/ManifestTypes";
import { useBoolean } from "@fluentui/react-hooks/lib/useBoolean";
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import ShimmerListComponent from './ShimmerListComponent';


const ShimmerButton: React.FC<ComponentFramework.Context<IInputs>> = (
  props
) => {
    const contentStyles = mergeStyleSets({
        container: {
          display: "flex",
          height:"clamp(500px,75vh,800px)",
          width:"clamp(500px,75vw,1200px)",
          padding:"1rem",
          flexFlow: "column nowrap",
          alignItems: "stretch",
        }
      });
    const [isModalOpen, { setTrue: showModal, setFalse: hideModal }] = useBoolean(false);
    const modalButtonClick = () => {
        showModal();
      };
    
  return (
      <>
        <Modal
          isOpen={isModalOpen}
          onDismiss={hideModal}
          isBlocking={false}
          containerClassName={contentStyles.container}
        >
            <ShimmerListComponent
                _context={props.webAPI}
            />
        </Modal>
        <DefaultButton text="Accounts List" onClick={modalButtonClick} />
      </>
  );
};
export default ShimmerButton;
