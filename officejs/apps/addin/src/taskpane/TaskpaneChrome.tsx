import { History, Settings, SquarePen } from "lucide-react";
import { FeatureIcon } from "./FeatureIcon";

export function TaskpaneChrome() {
  return (
    <>
      <header className="navbar agent-header">
        <div className="brand" hidden>
          <span className="sparkle" aria-hidden="true">
            <FeatureIcon id="surface-feature-icon" name="agent" size={16} />
          </span>
          <div className="brand-copy">
            <span className="product-edition">WordOllama.JS</span>
            <h1 id="surface-title">Agent</h1>
          </div>
        </div>
        <div
          id="runtime-status"
          className="runtime-strip"
          data-state="connecting"
          role="status"
          aria-live="polite"
        >
          <span className="runtime-indicator" aria-hidden="true" />
          <span
            id="runtime-status-text"
            className="runtime-strip-text"
            data-i18n="taskpane.runtime.connecting"
          >
            Connecting to local service…
          </span>
        </div>
        <div className="header-actions">
          <span
            id="host-status"
            className="status"
            data-i18n="taskpane.common.connectingWord"
          >
            Connecting to Word…
          </span>
          <button
            id="chat-history-toggle"
            className="btn btn-ghost btn-square btn-sm icon-button"
            type="button"
            aria-label="Chat history"
            title="Chat history"
            aria-expanded="false"
            aria-controls="chat-history-menu"
            data-i18n-aria-label="taskpane.agent.chatHistory"
            data-i18n-title="taskpane.agent.chatHistory"
          >
            <History size={15} strokeWidth={2} aria-hidden="true" />
          </button>
          <button
            id="new-chat"
            className="btn btn-ghost btn-square btn-sm icon-button"
            type="button"
            aria-label="New chat"
            title="New chat"
            data-i18n-aria-label="taskpane.agent.newChat"
            data-i18n-title="taskpane.agent.newChat"
          >
            <SquarePen size={15} strokeWidth={2} aria-hidden="true" />
          </button>
          <button
            id="open-settings"
            className="btn btn-ghost btn-square btn-sm icon-button"
            type="button"
            aria-label="Settings"
            title="Settings"
            data-i18n-aria-label="taskpane.common.settings"
            data-i18n-title="taskpane.common.settings"
          >
            <Settings size={15} strokeWidth={2} aria-hidden="true" />
          </button>
        </div>
        <div
          id="chat-history-menu"
          className="chat-history-menu"
          role="menu"
          aria-label="Chat history"
          data-i18n-aria-label="taskpane.agent.chatHistory"
          hidden
        />
      </header>
    </>
  );
}
