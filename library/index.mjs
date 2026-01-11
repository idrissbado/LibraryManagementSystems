// Example usage and dependency injection
default async function main() {
  const { Library } = await import('./library/library.js');
  const { InMemoryLibraryRepository } = await import('./library/repository.js');
  const { StandardLendingStrategy } = await import('./library/strategy.js');
  const { generateId } = await import('./library/utils.js');

  // Mock user
  const user = { id: 'u1', name: 'Alice', hasOverdueItems: false };

  // Dependency injection
  const repo = new InMemoryLibraryRepository();
  const strategy = new StandardLendingStrategy();
  const library = new Library(repo, strategy);

  // Observer example
  library.subscribe({
    update: (data) => console.log('Event:', data)
  });

  // Add items
  library.addItem('Book', { id: generateId(), title: '1984', author: 'George Orwell' });
  library.addItem('Magazine', { id: generateId(), title: 'Tech Monthly', issue: '2026-01' });

  // List items
  console.log('All items:', library.listItems());

  // Lend item
  const items = library.listItems();
  if (items.length > 0) {
    library.lendItem(user, items[0].id);
  }
}

main();
