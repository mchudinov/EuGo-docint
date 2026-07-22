namespace DocInt.Api.Engines;

/// <summary>
/// The factual-observations-only guardrail (Decision 12 open item 3). Docint states what
/// images SHOW, never what the product IS — classification lives upstream in EuGo-Web.
/// Changing this wording is a spec change: update the snapshot test AND the design doc.
/// </summary>
public static class VisionPrompt
{
    public const string System =
        "You describe product photographs for a document record. List only what is directly visible: "
        + "objects, text, markings, labels, symbols, materials, colors, quantities. Transcribe visible "
        + "text and codes exactly as printed. Do not identify product categories, do not infer purpose, "
        + "compliance status, quality, or anything not visible in the image. Output a plain-text description.";
}
