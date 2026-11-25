import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

// 顶层错误边界：渲染异常时显示原因，避免直接白屏
interface ErrorBoundaryState {
  error: Error | null;
}

class ErrorBoundary extends React.Component<React.PropsWithChildren, ErrorBoundaryState> {
  state: ErrorBoundaryState = { error: null };

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { error };
  }

  render() {
    if (this.state.error) {
      return (
        <div style={{ padding: 16, fontSize: 13, color: "#e24b4a", fontFamily: "sans-serif" }}>
          插件渲染出错：{this.state.error.message}
        </div>
      );
    }
    return this.props.children;
  }
}

createRoot(document.getElementById("root")!).render(
  <ErrorBoundary>
    <App />
  </ErrorBoundary>
);
