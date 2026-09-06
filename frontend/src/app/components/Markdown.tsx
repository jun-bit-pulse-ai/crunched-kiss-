import { Fragment } from "react";
import { parseMarkdown, type Inline } from "../markdown";

/**
 * Renders the small Markdown subset the model uses. Spans become React
 * elements, so model output is never injected as markup.
 */
export function Markdown({ text }: { text: string }) {
  const blocks = parseMarkdown(text);

  if (blocks.length === 0) {
    return null;
  }

  return (
    <>
      {blocks.map((block, index) => {
        switch (block.type) {
          case "heading":
            return (
              <h3 key={index} className="md-heading">
                <Spans spans={block.spans} />
              </h3>
            );
          case "list": {
            const items = block.items.map((item, itemIndex) => (
              <li key={itemIndex}>
                <Spans spans={item} />
              </li>
            ));
            return block.ordered ? (
              <ol key={index} className="md-list" start={block.start}>
                {items}
              </ol>
            ) : (
              <ul key={index} className="md-list">
                {items}
              </ul>
            );
          }
          case "table":
            return (
              // The pane is narrow, so a wide table scrolls inside its own box
              // instead of stretching the bubble.
              <div key={index} className="md-table-wrap">
                <table className="md-table">
                  <thead>
                    <tr>
                      {block.header.map((cell, cellIndex) => (
                        <th key={cellIndex}>
                          <Spans spans={cell} />
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {block.rows.map((row, rowIndex) => (
                      <tr key={rowIndex}>
                        {row.map((cell, cellIndex) => (
                          <td key={cellIndex}>
                            <Spans spans={cell} />
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            );
          case "rule":
            return <hr key={index} className="md-rule" />;
          default:
            return (
              <p key={index}>
                <Spans spans={block.spans} />
              </p>
            );
        }
      })}
    </>
  );
}

function Spans({ spans }: { spans: Inline[] }) {
  return (
    <>
      {spans.map((span, index) => {
        switch (span.type) {
          case "bold":
            return <strong key={index}>{span.text}</strong>;
          case "italic":
            return <em key={index}>{span.text}</em>;
          case "code":
            return (
              <code key={index} className="md-code">
                {span.text}
              </code>
            );
          default:
            return <Fragment key={index}>{span.text}</Fragment>;
        }
      })}
    </>
  );
}
