import { OfficeExtension } from "../api-extractor-inputs-office/office"
////////////////////////////////////////////////////////////////
///////////////////// Begin PowerPoint APIs ////////////////////
////////////////////////////////////////////////////////////////

export declare namespace PowerPoint {
    /**
     * Creates and opens a new presentation. Optionally, the presentation can be pre-populated with a base64-encoded .pptx file.
     *
     * [Api set: PowerPointApi 1.1]
     *
     * @param base64File - Optional. The base64-encoded .pptx file. The default value is null.
     */
    export function createPresentation(base64File?: string): Promise<void>;
}

////////////////////////////////////////////////////////////////
////////////////////// End PowerPoint APIs /////////////////////
////////////////////////////////////////////////////////////////