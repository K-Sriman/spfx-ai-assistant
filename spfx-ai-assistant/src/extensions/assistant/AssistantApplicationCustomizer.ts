import { override } from "@microsoft/decorators";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { initGraphClientFactory } from "../../services/graphClient";
import AssistantLauncher from "./components/AssistantLauncher";
import { trackInfo, trackError } from "../../utils/telemetry";

export interface IAssistantApplicationCustomizerProperties {
  testMessage?: string;
}

// Singleton host element — prevents duplicate mounts across SPA navigations
let _fabHost: HTMLElement | null = null;

/**
 * SPFx Application Customizer – AI Assistant (GPT-4o powered)
 *
 * Injects a floating assistant button + slide-in chat panel into every modern
 * SharePoint page. All intelligence is routed through Azure OpenAI GPT-4o.
 * All Graph/SP calls run in delegated user context — no app-only permissions.
 */
export default class AssistantApplicationCustomizer
  extends BaseApplicationCustomizer<IAssistantApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    trackInfo("AssistantApplicationCustomizer.onInit", {
      site: this.context.pageContext.web.absoluteUrl,
    });

    // Wire up Graph client factory (delegated)
    initGraphClientFactory(this.context.msGraphClientFactory);

    this._mountFab();

    // Re-mount on SPA navigation events
    this.context.placeholderProvider.changedEvent.add(this, this._mountFab.bind(this));

    return Promise.resolve();
  }

  @override
  public onDispose(): void {
    if (_fabHost) {
      ReactDOM.unmountComponentAtNode(_fabHost);
      _fabHost.parentElement?.removeChild(_fabHost);
      _fabHost = null;
    }
    super.onDispose();
  }

  private _mountFab(): void {
    try {
      if (!_fabHost) {
        _fabHost = document.createElement("div");
        _fabHost.id = "ai-assistant-root";
        document.body.appendChild(_fabHost);
      }

      ReactDOM.render(
        React.createElement(AssistantLauncher, { spContext: this.context }),
        _fabHost
      );
    } catch (err) {
      trackError("AssistantApplicationCustomizer._mountFab", err);
    }
  }
}
