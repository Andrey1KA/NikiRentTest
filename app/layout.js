import "./globals.css";

export const metadata = {
  title: "Обращения гостей — NikiRent",
  description: "Статистика обращений гостей и нерешённые бытовые вопросы",
};

export default function RootLayout({ children }) {
  return (
    <html lang="ru">
      <body>{children}</body>
    </html>
  );
}
