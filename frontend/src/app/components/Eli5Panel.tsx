type Eli5PanelProps = {
  text: string;
  onClose: () => void;
};

export function Eli5Panel({ text, onClose }: Eli5PanelProps) {
  return (
    <aside className="eli5-panel" aria-label="Explain like I'm 5">
      <div className="eli5-head">
        <p className="eli5-kicker">Like you are five</p>
        <button type="button" className="eli5-close" onClick={onClose}>
          Hide
        </button>
      </div>
      <p className="eli5-body">{text}</p>
    </aside>
  );
}
