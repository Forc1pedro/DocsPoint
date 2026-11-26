'use client';

import React, { useState } from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Input,
  Button,
  Label,
  Spinner,
} from '@fluentui/react-components';
import {
  FolderOpen24Regular,
  LockClosed20Regular,
  Mail20Regular,
} from '@fluentui/react-icons';
import { api } from '../../services/api';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    minHeight: '100vh',
    width: '100vw',
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.padding('24px'),
  },
  card: {
    width: '100%',
    maxWidth: '440px',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius('8px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    boxShadow: '0 8px 24px rgba(0,0,0,0.08)',
    ...shorthands.padding('40px', '36px'),
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('24px'),
  },
  header: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    textAlign: 'center',
    ...shorthands.gap('12px'),
  },
  logoBox: {
    width: '48px',
    height: '48px',
    backgroundColor: '#0078d4',
    ...shorthands.borderRadius('8px'),
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: '#ffffff',
  },
  title: {
    fontSize: '22px',
    fontWeight: '600',
    color: tokens.colorNeutralForeground1,
  },
  subtitle: {
    fontSize: '13px',
    color: tokens.colorNeutralForeground3,
    marginTop: '2px',
  },
  form: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('16px'),
  },
  fieldGroup: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('6px'),
  },
  errorBox: {
    backgroundColor: 'rgba(209, 52, 56, 0.1)',
    color: '#d13438',
    ...shorthands.padding('10px', '14px'),
    ...shorthands.borderRadius('4px'),
    fontSize: '13px',
    ...shorthands.border('1px', 'solid', 'rgba(209, 52, 56, 0.3)'),
  },
});

export default function LoginPage() {
  const styles = useStyles();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [errorMessage, setErrorMessage] = useState('');
  const [isLoading, setIsLoading] = useState(false);

  const handleStandardLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!email.trim() || !password.trim()) {
      setErrorMessage('Informe seu e-mail corporativo e sua senha.');
      return;
    }
    setErrorMessage('');
    setIsLoading(true);
    try {
      await api.login(email.trim(), password);
      window.location.href = '/';
    } catch (err: any) {
      setErrorMessage(
        err.response?.data?.error || 'Falha ao autenticar. Verifique suas credenciais.'
      );
      setIsLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.card}>
        <div className={styles.header}>
          <div className={styles.logoBox}>
            <FolderOpen24Regular />
          </div>
          <div>
            <div className={styles.title}>Microsoft SharePoint</div>
            <div className={styles.subtitle}>Portal Corporativo e Colaboração Segura</div>
          </div>
        </div>

        {errorMessage && <div className={styles.errorBox}>{errorMessage}</div>}

        <form onSubmit={handleStandardLogin} className={styles.form}>
          <div className={styles.fieldGroup}>
            <Label htmlFor="login-email">E-mail corporativo:</Label>
            <Input
              id="login-email"
              data-testid="login-email-input"
              type="email"
              placeholder="nome@organizacao.com"
              value={email}
              onChange={(_, d) => setEmail(d.value)}
              contentBefore={<Mail20Regular />}
              disabled={isLoading}
              autoFocus
            />
          </div>

          <div className={styles.fieldGroup}>
            <Label htmlFor="login-password">Senha:</Label>
            <Input
              id="login-password"
              data-testid="login-password-input"
              type="password"
              placeholder="••••••••"
              value={password}
              onChange={(_, d) => setPassword(d.value)}
              contentBefore={<LockClosed20Regular />}
              disabled={isLoading}
            />
          </div>

          <Button
            type="submit"
            appearance="primary"
            data-testid="login-submit-btn"
            disabled={isLoading}
            style={{ width: '100%', marginTop: '8px' }}
          >
            {isLoading ? <Spinner size="tiny" label="Entrando..." /> : 'Entrar'}
          </Button>
        </form>
      </div>
    </div>
  );
}
