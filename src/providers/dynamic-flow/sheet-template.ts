export interface SheetTemplate {
    id: string;
    import: any;
  }
  
  export class SheetTemplateProvider {
    private templates: SheetTemplate[] = [
      {
        id: "templateId",
        import: {
          sheets: ["Sheet1"],
          columnToKey: { A: "name", B: "age", C: "email" },
          data: { startRow: 2 },
        },
      },
    ];
  
    async getOne(query: { where: { id: string } }): Promise<SheetTemplate | null> {
      const { id } = query.where;
      const template = this.templates.find((template) => template.id === id);
      return template || null;
    }
  }
  