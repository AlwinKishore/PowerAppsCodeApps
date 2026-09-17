export function ActionButtons({ onReset, onNewRequest }: { onReset: () => void; onNewRequest: () => void }) {
  return <div className="actions">
          <button onClick={onReset}>⟳ &nbsp;Refresh</button>
          <button className="new-request" onClick={onNewRequest}>＋ &nbsp;New Request</button>
        </div>
}
