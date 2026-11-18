import { useEffect, useState } from 'react';
import { api, type Me } from './api';
import { AlertBanner } from './components/AlertBanner';
import { Icon, type IconName } from './components/Icon';
import { BillingTile } from './components/tiles/BillingTile';
import { OnboardingTile } from './components/tiles/OnboardingTile';
import { ReferralsTile } from './components/tiles/ReferralsTile';
import { SetupTile } from './components/tiles/SetupTile';
import { ReflectiveSupervisionTile } from './components/tiles/ReflectiveSupervisionTile';
import { StaffExpectationsTile } from './components/tiles/StaffExpectationsTile';
import { StaffInformationTile } from './components/tiles/StaffInformationTile';
import { AnnouncementPanel } from './components/AnnouncementPanel';
import { StaffConnection } from './components/StaffConnection';

type Color = 'green' | 'pink' | 'yellow' | 'blue';

const TILES: Array<{ id: string; href: string; title: string; text: string; icon: IconName; color: Color }> = [
  { id: 'billing', href: '#/billing', title: 'Billing', text: 'Invoice status from IntakeQ', icon: 'receipt', color: 'green' },
  { id: 'onboarding', href: '#/onboarding', title: 'Onboarding', text: 'Checklists and employee documents', icon: 'user-plus', color: 'pink' },
  { id: 'referrals', href: '#/referrals', title: 'Referrals', text: 'Incoming referrals', icon: 'share', color: 'blue' },
  { id: 'staff', href: '#/staff', title: 'Staff Information', text: 'Trainings and open notes', icon: 'users', color: 'yellow' },
  { id: 'expectations', href: '#/expectations', title: 'Staff Expectations', text: 'Openings, mileage, and compliance', icon: 'calendar-check', color: 'green' },
  { id: 'reflective-supervision', href: '#/reflective-supervision', title: 'Reflective Supervision', text: 'Observation process files', icon: 'messages', color: 'pink' },
  { id: 'setup', href: '#/setup', title: 'Setup', text: 'SharePoint site, lists, and folders', icon: 'settings', color: 'blue' },
];

function useHash(): string {
  const [hash, setHash] = useState(window.location.hash || '#/');
  useEffect(() => {
    const onChange = () => setHash(window.location.hash || '#/');
    window.addEventListener('hashchange', onChange);
    return () => window.removeEventListener('hashchange', onChange);
  }, []);
  return hash;
}

function viewFor(hash: string): string {
  const path = hash.replace(/^#/, '');
  const match = TILES.find((tile) => path === '/' + tile.id || path.startsWith('/' + tile.id + '/'));
  return match ? match.id : 'home';
}

function greetingFor(name: string): string {
  const hour = Number(new Intl.DateTimeFormat('en-US', { hour: 'numeric', hour12: false, timeZone: 'America/Denver' }).format(new Date())) % 24;
  const part = hour < 12 ? 'Good morning' : hour < 17 ? 'Good afternoon' : 'Good evening';
  const first = name.includes('@') ? '' : name.trim().split(/\s+/)[0];
  return first ? part + ', ' + first : part;
}

export function App() {
  const hash = useHash();
  const view = viewFor(hash);
  const [me, setMe] = useState<Me | null>(null);
  const [error, setError] = useState('');

  useEffect(() => {
    api<Me>('/api/me')
      .then((data) => setMe(data))
      .catch((err: Error) => setError(err.message));
  }, []);

  const page = view === 'billing' ? <BillingTile />
    : view === 'onboarding' ? <OnboardingTile />
    : view === 'referrals' ? <ReferralsTile />
    : view === 'staff' ? <StaffInformationTile />
    : view === 'expectations' ? <StaffExpectationsTile />
    : view === 'reflective-supervision' ? <ReflectiveSupervisionTile />
    : view === 'setup' ? <SetupTile />
    : null;

  const today = new Intl.DateTimeFormat('en-US', {
    weekday: 'long', month: 'long', day: 'numeric', year: 'numeric', timeZone: 'America/Denver',
  }).format(new Date());

  return (
    <>
      <header className="site-header" role="banner">
        <a className="site-header__logo" href="/">Emme</a>
        <div className="site-header__actions">
          <time className="site-header__date">{today}</time>
          <span className="site-header__date">{me?.name || me?.email || ''}</span>
          <a className="btn btn--outline-light" href="/">Provider dashboard</a>
        </div>
      </header>

      <nav className="tab-nav" aria-label="Admin navigation">
        <a className={view === 'home' ? 'tab-nav__btn tab-nav__btn--active' : 'tab-nav__btn'} href="#/" aria-current={view === 'home' ? 'page' : undefined}>
          <Icon name="dashboard" className="icon-sm" />
          <span>Admin Home</span>
        </a>
        {TILES.map((tile) => (
          <a
            key={tile.id}
            className={view === tile.id ? 'tab-nav__btn tab-nav__btn--active' : 'tab-nav__btn'}
            href={tile.href}
            aria-current={view === tile.id ? 'page' : undefined}
          >
            <Icon name={tile.icon} className="icon-sm" />
            <span>{tile.title}</span>
          </a>
        ))}
      </nav>

      <div className="content-wrapper">
        <AlertBanner />
        {error ? <p className="error">{error}</p> : null}
        {me && !me.isAdmin ? <p className="error">Emme Admin access is required for these tools.</p> : null}

        {me?.isAdmin && view === 'home' ? (
          <>
            <section className="greeting" aria-label="Greeting">
              <h1 className="greeting__title">{greetingFor(me.name || me.email || '')}</h1>
              <p className="greeting__quote">Everything you need to run the practice, in one place.</p>
            </section>

            <section className="announce" aria-label="Staff announcement">
              <AnnouncementPanel />
            </section>

            <section aria-labelledby="admin-tools-heading">
              <h2 className="section-heading" id="admin-tools-heading">Admin tools</h2>
              <ul className="card-grid">
                {TILES.map((tile) => (
                  <li key={tile.id}>
                    <a className={'card card--' + tile.color} href={tile.href}>
                      <div className={'card__icon-wrap card__icon-wrap--' + tile.color}>
                        <Icon name={tile.icon} className="icon-md" />
                      </div>
                      <div className="card__body">
                        <h3 className="card__title">{tile.title}</h3>
                        <p className="card__desc">{tile.text}</p>
                      </div>
                      <span className={'card__action btn btn--' + tile.color} aria-hidden="true">Open →</span>
                    </a>
                  </li>
                ))}
              </ul>
            </section>
          </>
        ) : null}

        {me?.isAdmin && page ? <div className="widget">{page}</div> : null}
      </div>

      <footer className="site-footer" role="contentinfo">
        <p className="site-footer__copy">© 2026 Emme Therapy · Admin dashboard</p>
      </footer>

      {me?.isAdmin ? <StaffConnection /> : null}
    </>
  );
}
