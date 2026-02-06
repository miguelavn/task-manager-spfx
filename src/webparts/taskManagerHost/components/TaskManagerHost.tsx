import * as React from 'react';
import styles from './TaskManagerHost.module.scss';
import type { ITaskManagerHostProps } from './ITaskManagerHostProps';
import type { ITask } from './ITask';
import { SharePointService } from '../services/SharePointService';

interface ITaskManagerHostState {
  tasks: ITask[];
  newTask: ITask;
  editingTaskId: number | undefined;
  isLoading: boolean;
  error: string | undefined;
  isFullScreen: boolean;
}

export default class TaskManagerHost extends React.Component<ITaskManagerHostProps, ITaskManagerHostState> {
  private readonly _sharePointService: SharePointService;
  private readonly _hideChromeBodyClassName: string = 'taskManagerHideSharePointChrome';
  private readonly _hideChromeStyleElementId: string = 'taskManagerHideSharePointChromeStyle';
  private _isSubmitting: boolean = false;

  public constructor(props: ITaskManagerHostProps) {
    super(props);

    this._sharePointService = new SharePointService(props.context);

    this.state = {
      tasks: [],
      newTask: { Title: '', Responsable: '', Estado: 'Pendiente' },
      editingTaskId: undefined,
      isLoading: false,
      error: undefined,
      isFullScreen: true
    };
  }

  public async componentDidMount(): Promise<void> {
    this._ensureHideChromeStyle();
    this._setSharePointChromeHidden(this.state.isFullScreen);
    await this._loadTasks();
  }

  public componentWillUnmount(): void {
    // Restaurar el chrome al salir de la página (o si el webpart se desmonta)
    this._setSharePointChromeHidden(false);
  }

  public componentDidUpdate(prevProps: ITaskManagerHostProps, prevState: ITaskManagerHostState): void {
    if (prevState.isFullScreen !== this.state.isFullScreen) {
      this._setSharePointChromeHidden(this.state.isFullScreen);
    }
  }

  private _ensureHideChromeStyle(): void {
    if (!document?.head) return;
    if (document.getElementById(this._hideChromeStyleElementId)) return;

    const style = document.createElement('style');
    style.id = this._hideChromeStyleElementId;
    style.type = 'text/css';
    style.textContent = `
      body.${this._hideChromeBodyClassName} #SuiteNavWrapper,
      body.${this._hideChromeBodyClassName} div[data-automation-id='SuiteNav'],
      body.${this._hideChromeBodyClassName} #spSiteHeader,
      body.${this._hideChromeBodyClassName} div[data-automation-id='SiteHeader'] {
        display: none !important;
      }

      body.${this._hideChromeBodyClassName} {
        --spfx-taskmanager-hide-chrome: 1;
      }
    `;

    document.head.appendChild(style);
  }

  private _setSharePointChromeHidden(hidden: boolean): void {
    if (!document?.body) return;
    // Nota: esto es un hack visual (no soportado oficialmente por Microsoft) y depende del DOM de SharePoint.
    // Se aplica solo en páginas donde este webpart está presente.
    if (hidden) {
      document.body.classList.add(this._hideChromeBodyClassName);
    } else {
      document.body.classList.remove(this._hideChromeBodyClassName);
    }
  }

  private async _loadTasks(): Promise<void> {
    this.setState({ isLoading: true, error: undefined });
    try {
      const tasks = await this._sharePointService.getTasks();
      this.setState({ tasks });
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error('Error cargando tareas:', error);
      const message = error instanceof Error ? error.message : 'Error cargando tareas';
      this.setState({ error: message });
    } finally {
      this.setState({ isLoading: false });
    }
  }

  private _handleInputChange(field: keyof ITask, value: string): void {
    this.setState(prevState => ({
      newTask: {
        ...prevState.newTask,
        [field]: value
      }
    }));
  }

  private async _handleCreateTask(e: React.FormEvent): Promise<void> {
    e.preventDefault();

    if (this._isSubmitting) {
      return;
    }

    if (!this.state.newTask.Title.trim() || !this.state.newTask.Responsable.trim()) {
      this.setState({ error: 'Por favor completa todos los campos' });
      return;
    }

    this._isSubmitting = true;
    this.setState({ isLoading: true });
    try {
      await this._sharePointService.createTask(this.state.newTask);
      this.setState({ newTask: { Title: '', Responsable: '', Estado: 'Pendiente' }, error: undefined });
      await this._loadTasks();
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error('Error creando tarea:', error);
      const message = error instanceof Error ? error.message : 'Error creando tarea';
      this.setState({ error: message });
    } finally {
      this._isSubmitting = false;
      this.setState({ isLoading: false });
    }
  }

  private async _handleUpdateTask(e: React.FormEvent): Promise<void> {
    e.preventDefault();

    if (this._isSubmitting) {
      return;
    }

    if (!this.state.editingTaskId) return;

    this._isSubmitting = true;
    this.setState({ isLoading: true });
    try {
      await this._sharePointService.updateTask(this.state.editingTaskId, this.state.newTask);
      this.setState({
        newTask: { Title: '', Responsable: '', Estado: 'Pendiente' },
        editingTaskId: undefined,
        error: undefined
      });
      await this._loadTasks();
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error('Error actualizando tarea:', error);
      const message = error instanceof Error ? error.message : 'Error actualizando tarea';
      this.setState({ error: message });
    } finally {
      this._isSubmitting = false;
      this.setState({ isLoading: false });
    }
  }

  private _handleEditTask(task: ITask): void {
    this.setState({
      editingTaskId: task.ID,
      newTask: {
        Title: task.Title,
        Responsable: task.Responsable,
        Estado: task.Estado
      }
    });
  }

  private _handleCancelEdit(): void {
    this.setState({
      editingTaskId: undefined,
      newTask: { Title: '', Responsable: '', Estado: 'Pendiente' }
    });
  }

  private async _handleDeleteTask(taskId: number | undefined): Promise<void> {
    if (!taskId) return;

    if (this._isSubmitting) {
      return;
    }

    // eslint-disable-next-line no-alert
    if (window.confirm('¿Estás seguro de que quieres eliminar esta tarea?')) {
      this._isSubmitting = true;
      this.setState({ isLoading: true });
      try {
        await this._sharePointService.deleteTask(taskId);
        await this._loadTasks();
      } catch (error) {
        // eslint-disable-next-line no-console
        console.error('Error eliminando tarea:', error);
        const message = error instanceof Error ? error.message : 'Error eliminando tarea';
        this.setState({ error: message });
      } finally {
        this._isSubmitting = false;
        this.setState({ isLoading: false });
      }
    }
  }

  private _handleToggleFullScreen(): void {
    this.setState(prevState => ({
      isFullScreen: !prevState.isFullScreen
    }));
  }

  private _getContainerClassName(): string {
    return this.state.isFullScreen
      ? styles.fullScreenContainer
      : styles.container;
  }

  private _getStatusClassName(estado: string): string {
    switch ((estado || '').trim().toLowerCase()) {
      case 'pendiente':
        return styles.statusPendiente;
      case 'en progreso':
        return styles.statusEnProgreso;
      case 'completada':
        return styles.statusCompletada;
      case 'cancelada':
        return styles.statusCancelada;
      default:
        return styles.statusPendiente;
    }
  }

  public render(): React.ReactElement<ITaskManagerHostProps> {
    const { tasks, newTask, editingTaskId, isLoading, error, isFullScreen } = this.state;

    return (
      <div className={styles.taskManagerHost}>
        <div className={this._getContainerClassName()}>
          <div className={styles.header}>
            <h1 className={styles.title}>Gestor de Tareas</h1>
            <button
              className={styles.fullScreenBtn}
              onClick={() => this._handleToggleFullScreen()}
              title={isFullScreen ? 'Salir de pantalla completa' : 'Pantalla completa'}
            >
              {isFullScreen ? 'Salir' : 'Pantalla completa'}
            </button>
          </div>

          {error && (
            <div className={styles.error}>
              <span>{error}</span>
              <button
                className={styles.closeErrorBtn}
                onClick={() => this.setState({ error: undefined })}
              >
                Cerrar
              </button>
            </div>
          )}

          <form
            className={styles.form}
            onSubmit={editingTaskId
              ? (e) => this._handleUpdateTask(e)
              : (e) => this._handleCreateTask(e)
            }
          >
            <h2>{editingTaskId ? 'Editar Tarea' : 'Nueva Tarea'}</h2>

            <div className={styles.formGroup}>
              <label htmlFor="taskTitle">Título de la Tarea *</label>
              <input
                id="taskTitle"
                type="text"
                value={newTask.Title}
                onChange={(e) => this._handleInputChange('Title', e.target.value)}
                placeholder="Ingresa el título de la tarea"
                required
              />
            </div>

            <div className={styles.formGroup}>
              <label htmlFor="taskResponsable">Responsable *</label>
              <input
                id="taskResponsable"
                type="text"
                value={newTask.Responsable}
                onChange={(e) => this._handleInputChange('Responsable', e.target.value)}
                placeholder="Nombre del responsable"
                required
              />
            </div>

            <div className={styles.formGroup}>
              <label htmlFor="taskEstado">Estado *</label>
              <select
                id="taskEstado"
                value={newTask.Estado}
                onChange={(e) => this._handleInputChange('Estado', e.target.value)}
              >
                <option value="Pendiente">Pendiente</option>
                <option value="En Progreso">En Progreso</option>
                <option value="Completada">Completada</option>
                <option value="Cancelada">Cancelada</option>
              </select>
            </div>

            <div className={styles.buttonGroup}>
              <button type="submit" className={styles.submitBtn} disabled={isLoading}>
                {isLoading ? 'Guardando...' : (editingTaskId ? 'Actualizar' : 'Crear Tarea')}
              </button>
              {editingTaskId && (
                <button
                  type="button"
                  className={styles.cancelBtn}
                  onClick={() => this._handleCancelEdit()}
                  disabled={isLoading}
                >
                  Cancelar
                </button>
              )}
            </div>
          </form>

          <div className={styles.tasksSection}>
            <h2>Tareas ({tasks.length})</h2>

            {isLoading && tasks.length === 0 && (
              <div className={styles.loading}>Cargando tareas...</div>
            )}

            {tasks.length === 0 && !isLoading && (
              <div className={styles.noTasks}>
                No hay tareas creadas. ¡Crea una nueva!
              </div>
            )}

            <div className={styles.tasksList}>
              {tasks.map((task) => (
                <div key={task.ID} className={styles.taskCard}>
                  <div className={styles.taskHeader}>
                    <h3 className={styles.taskTitle}>{task.Title}</h3>
                    <span className={`${styles.status} ${this._getStatusClassName(task.Estado)}`}>
                      {task.Estado}
                    </span>
                  </div>

                  <div className={styles.taskInfo}>
                    <div className={styles.infoRow}>
                      <span className={styles.label}>Responsable:</span>
                      <span className={styles.value}>{task.Responsable}</span>
                    </div>
                    {task.Autor && (
                      <div className={styles.infoRow}>
                        <span className={styles.label}>Autor:</span>
                        <span className={styles.value}>{task.Autor}</span>
                      </div>
                    )}
                    {task.Creado && (
                      <div className={styles.infoRow}>
                        <span className={styles.label}>Creada:</span>
                        <span className={styles.value}>
                          {new Date(task.Creado).toLocaleDateString('es-ES')}
                        </span>
                      </div>
                    )}
                  </div>

                  <div className={styles.taskActions}>
                    <button
                      className={styles.editBtn}
                      onClick={() => this._handleEditTask(task)}
                      disabled={isLoading}
                    >
                      Editar
                    </button>
                    <button
                      className={styles.deleteBtn}
                      onClick={() => this._handleDeleteTask(task.ID)}
                      disabled={isLoading}
                    >
                      Eliminar
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
