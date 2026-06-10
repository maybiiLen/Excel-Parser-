import { PasteInput } from "@/components/PasteInput";

export default function Home() {
  return (
    <main className="mx-auto flex w-full max-w-3xl flex-1 flex-col gap-8 px-6 py-12 sm:py-16">
      <header className="flex flex-col gap-2">
        <h1 className="text-2xl font-bold tracking-tight sm:text-3xl">
          Excel &rarr; Word Sections
        </h1>
        <p className="text-foreground/60">
          Paste a table copied from Excel or Google Sheets to preview it as
          numbered Word-style sections. Adjust the starting number, or toggle to
          inspect the raw parsed JSON.
        </p>
      </header>
      <PasteInput />
    </main>
  );
}
