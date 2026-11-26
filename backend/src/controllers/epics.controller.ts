import { Request, Response } from 'express';
import { listService } from '../services/list.service';
import { pageService } from '../services/page.service';
import { permissionService } from '../services/permission.service';

export class EpicsController {
  // --------------------------------------------------------------------------
  // ÉPICO 5: LISTS
  // --------------------------------------------------------------------------
  public async getLists(req: Request, res: Response): Promise<void> {
    try {
      const siteId = req.params.siteId || '11111111-1111-1111-1111-111111111111';
      const lists = await listService.getLists(siteId);
      res.json(lists);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async createList(req: Request, res: Response): Promise<void> {
    try {
      const siteId = req.params.siteId || '11111111-1111-1111-1111-111111111111';
      const { title, description } = req.body;
      if (!title) {
        res.status(400).json({ error: 'O título da lista é obrigatório' });
        return;
      }
      const list = await listService.createList(siteId, title, description);
      res.status(201).json(list);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async getList(req: Request, res: Response): Promise<void> {
    try {
      const list = await listService.getList(req.params.listId);
      if (!list) {
        res.status(404).json({ error: 'Lista não encontrada' });
        return;
      }
      res.json(list);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async getColumns(req: Request, res: Response): Promise<void> {
    try {
      const columns = await listService.getColumns(req.params.listId);
      res.json(columns);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async addColumn(req: Request, res: Response): Promise<void> {
    try {
      const { name, displayName, columnType, options, isRequired } = req.body;
      if (!name || !displayName || !columnType) {
        res.status(400).json({ error: 'Campos name, displayName e columnType são obrigatórios' });
        return;
      }
      const column = await listService.addColumn(
        req.params.listId,
        name,
        displayName,
        columnType,
        options,
        isRequired
      );
      res.status(201).json(column);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async getItems(req: Request, res: Response): Promise<void> {
    try {
      const items = await listService.getItems(req.params.listId);
      res.json(items);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async createItem(req: Request, res: Response): Promise<void> {
    try {
      const { data } = req.body;
      const userId = (req as any).user?.id;
      const item = await listService.createItem(req.params.listId, data, userId);
      res.status(201).json(item);
    } catch (err: any) {
      const status = err.status || 500;
      res.status(status).json({ error: err.message });
    }
  }

  public async updateItem(req: Request, res: Response): Promise<void> {
    try {
      const { data } = req.body;
      const userId = (req as any).user?.id;
      const item = await listService.updateItem(req.params.itemId, data, userId);
      res.json(item);
    } catch (err: any) {
      const status = err.status || 500;
      res.status(status).json({ error: err.message });
    }
  }

  public async deleteItem(req: Request, res: Response): Promise<void> {
    try {
      await listService.deleteItem(req.params.itemId);
      res.status(204).send();
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // ÉPICO 6: SITE PAGES & WEB PARTS CANVAS
  // --------------------------------------------------------------------------
  public async getPages(req: Request, res: Response): Promise<void> {
    try {
      const siteId = req.params.siteId || '11111111-1111-1111-1111-111111111111';
      const pages = await pageService.getSitePages(siteId);
      res.json(pages);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async getPage(req: Request, res: Response): Promise<void> {
    try {
      const page = await pageService.getPageById(req.params.pageId);
      if (!page) {
        res.status(404).json({ error: 'Página não encontrada' });
        return;
      }
      res.json(page);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async createPage(req: Request, res: Response): Promise<void> {
    try {
      const siteId = req.params.siteId || '11111111-1111-1111-1111-111111111111';
      const { title, description, canvasContent } = req.body;
      if (!title) {
        res.status(400).json({ error: 'O título da página é obrigatório' });
        return;
      }
      const page = await pageService.createPage(
        siteId,
        title,
        description,
        canvasContent,
        (req as any).user?.name || 'Carlos Silva'
      );
      res.status(201).json(page);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async updatePage(req: Request, res: Response): Promise<void> {
    try {
      const { title, canvasContent, isPublished } = req.body;
      const page = await pageService.updatePage(
        req.params.pageId,
        title,
        canvasContent,
        isPublished
      );
      res.json(page);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async deletePage(req: Request, res: Response): Promise<void> {
    try {
      await pageService.deletePage(req.params.pageId);
      res.status(204).send();
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // ÉPICO 7: GOVERNANÇA, PERMISSÕES E HERANÇA
  // --------------------------------------------------------------------------
  public async getItemPermissions(req: Request, res: Response): Promise<void> {
    try {
      const perms = await permissionService.getItemPermissions(req.params.itemId);
      res.json(perms);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async breakInheritance(req: Request, res: Response): Promise<void> {
    try {
      const { userEmail, permission } = req.body;
      await permissionService.breakInheritance(req.params.itemId, userEmail, permission);
      res.json({ success: true, message: 'Herança de permissões interrompida com sucesso.' });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async restoreInheritance(req: Request, res: Response): Promise<void> {
    try {
      await permissionService.restoreInheritance(req.params.itemId);
      res.json({ success: true, message: 'Herança de permissões restaurada com sucesso.' });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async setItemPermission(req: Request, res: Response): Promise<void> {
    try {
      const { userEmail, permission } = req.body;
      if (!userEmail || !permission) {
        res.status(400).json({ error: 'userEmail e permission são obrigatórios' });
        return;
      }
      const perm = await permissionService.setItemPermission(
        req.params.itemId,
        userEmail,
        permission
      );
      res.status(201).json(perm);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async removeUserPermission(req: Request, res: Response): Promise<void> {
    try {
      const { userEmail } = req.body;
      await permissionService.removeUserPermission(req.params.itemId, userEmail);
      res.status(204).send();
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async getSiteMembers(req: Request, res: Response): Promise<void> {
    try {
      const siteId = req.params.siteId || '11111111-1111-1111-1111-111111111111';
      const members = await permissionService.getSiteMembers(siteId);
      res.json(members);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async addSiteMember(req: Request, res: Response): Promise<void> {
    try {
      const siteId = req.params.siteId || '11111111-1111-1111-1111-111111111111';
      const { email, name, role } = req.body;
      if (!email || !name || !role) {
        res.status(400).json({ error: 'email, name e role são obrigatórios' });
        return;
      }
      const member = await permissionService.addSiteMember(siteId, email, name, role);
      res.status(201).json(member);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }
}

export const epicsController = new EpicsController();
