// Factory pattern for creating library items
export class Book {
  constructor({ id, title, author }) {
    this.id = id;
    this.title = title;
    this.author = author;
    this.type = 'Book';
  }
}

export class Magazine {
  constructor({ id, title, issue }) {
    this.id = id;
    this.title = title;
    this.issue = issue;
    this.type = 'Magazine';
  }
}

export class LibraryItemFactory {
  static createItem(type, props) {
    switch (type) {
      case 'Book':
        return new Book(props);
      case 'Magazine':
        return new Magazine(props);
      default:
        throw new Error('Unknown item type');
    }
  }
}
