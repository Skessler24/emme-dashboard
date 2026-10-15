import { useEffect, useState } from 'react';
import { api, type Me } from './api';
import { AlertBanner } from './components/AlertBanner';
import { BillingTile } from './components/tiles/BillingTile';
import { OnboardingTile } from './components/tiles/OnboardingTile';
import { ReferralsTile } from './components/tiles/ReferralsTile';
import { StaffExpectationsTile } from './components/tiles/StaffExpectationsTile';
import { StaffInformationTile } from './components/tiles/StaffInformationTile';

const TILES = [
  { id: 'billing', href: '#/billing', title: 'Billing', text: 'Invoice status from IntakeQ' },
  { id: 'onboarding', href: '#/onboarding', title: 'Onboarding', text: 'Checklists and employee documents' },
  { id: 'referrals', href: '#/referrals', title: 'Referrals', text: 'Incoming referrals' },
  { id: 'staff', href: '#/staff', title: 'Staff Information', text: 'Trainings, notes, and questions' },
  { id: 'expectations', href: '#/expectations', title: 'Staff Expectations', text: 'Openings, mileage, and compliance' },
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
    : null;

  return (
    <div className="shell">
      <header className="top">
        <a className="brand" href="/">Emme</a>
        <span className="who">{me?.name || me?.email || ''}</span>
      </header>
      <AlertBanner />
      {error ? <p className="error">{error}</p> : null}
      {me && !me.isAdmin ? <p className="error">Emme Admin access is required for these tools.</p> : null}
      <nav className="tiles" aria-label="Admin tools">
        {TILES.map((tile) => (
          <a key={tile.id} className={view === tile.id ? 'tile tile--active' : 'tile'} href={tile.href}>
            <strong>{tile.title}</strong>
            <span>{tile.text}</span>
          </a>
        ))}
      </nav>
      {me?.isAdmin && page ? <div className="panel">{page}</div> : null}
      {me?.isAdmin && !page ? <p className="lede">Choose a tile. Scheduling openings are on Staff Expectations.</p> : null}
    </div>
  );
}
