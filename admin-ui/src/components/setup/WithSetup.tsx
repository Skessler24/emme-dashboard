import type { ReactNode } from 'react';
import { SetupPanel } from './SetupPanel';

/** The marker that turns a tile into its setup questions. It starts with ~ so it can never match a SharePoint folder name in the web address. */
export const SETUP_MARK = '~setup';

export function onSetupTab(): boolean {
  return window.location.hash.replace(/\/+$/, '').endsWith('/' + SETUP_MARK);
}

/**
 * Adds a second tab, "Setup questions", to a tile. The tile's own screen is
 * passed in unchanged as the children and is only shown on the first tab.
 */
export function WithSetup(props: { topicId: string; mainLabel: string; children: ReactNode }) {
  const onSetup = onSetupTab();
  const base = '#/' + props.topicId;
  return (
    <section>
      <nav className="subtabs" aria-label="Sections">
        <a className={onSetup ? 'subtab' : 'subtab subtab--active'} href={base}>{props.mainLabel}</a>
        <a className={onSetup ? 'subtab subtab--active' : 'subtab'} href={base + '/' + SETUP_MARK}>Setup questions</a>
      </nav>
      {onSetup ? <SetupPanel topicId={props.topicId} /> : props.children}
    </section>
  );
}
