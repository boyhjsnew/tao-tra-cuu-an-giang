import { render, screen } from '@testing-library/react';
import App from './App';

test('renders import excel title', () => {
  render(<App />);
  const titleElement = screen.getByText(/Import Excel/i);
  expect(titleElement).toBeInTheDocument();
});
