import React from "react";

interface ErrorBoundaryProps {
  /** Human label shown in the fallback UI (e.g. "Dashboard", "Project card"). */
  scope?: string;
  /** Render prop for a custom fallback. Receives the caught error and a retry fn. */
  fallback?: (error: Error, retry: () => void) => React.ReactNode;
  /** If true, a caught error collapses the subtree silently to a small inline chip. */
  inline?: boolean;
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  error: Error | null;
}

export default class ErrorBoundary extends React.Component<
  ErrorBoundaryProps,
  ErrorBoundaryState
> {
  state: ErrorBoundaryState = { error: null };

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { error };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo): void {
    const scope = this.props.scope || "unknown";
    console.warn(`[ErrorBoundary:${scope}]`, error, info.componentStack);
  }

  retry = (): void => {
    this.setState({ error: null });
  };

  render(): React.ReactNode {
    const { error } = this.state;
    if (!error) return this.props.children;
    if (this.props.fallback) return this.props.fallback(error, this.retry);
    if (this.props.inline) {
      return (
        <div className="errorBoundaryInline" role="alert">
          <span>Could not display this item.</span>
          <button type="button" onClick={this.retry}>
            Retry
          </button>
        </div>
      );
    }
    return (
      <div className="errorBoundaryFallback" role="alert">
        <div className="errorBoundaryTitle">Something went wrong.</div>
        <div className="errorBoundaryMessage">
          {this.props.scope ? `The ${this.props.scope} hit an error. ` : ""}
          Your data is saved — reload the page or try again.
        </div>
        <div className="errorBoundaryDetails">{error.message}</div>
        <div className="errorBoundaryActions">
          <button
            type="button"
            className="errorBoundaryRetry"
            onClick={this.retry}
          >
            Try again
          </button>
          <button
            type="button"
            className="errorBoundaryReload"
            onClick={() => window.location.reload()}
          >
            Reload
          </button>
        </div>
      </div>
    );
  }
}
