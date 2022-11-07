import * as React from 'react';
import styles from './PersonalDashboard.module.scss';
import { IPersonalDashboardProps } from './IPersonalDashboardProps';
import { Agenda, File, Tasks, Todo } from '@microsoft/mgt-react/dist/es6/spfx';
import { DefaultButton, Label, PrimaryButton, Stack, StackItem, Text, TextField, TooltipHost } from 'office-ui-fabric-react';
import useLinks from '../hooks/useLinks';
import { IUserLink } from '../interfaces/IUserLink';
import useTrending from '../hooks/useTrending';
import { IUserTrending } from '../interfaces/IUserTrending';

const PersonalDashboard: React.FC<IPersonalDashboardProps> = ({
  context,
}) => {
  const [date, setDate] = React.useState<Date>(new Date());
  const [createLink, setCreateLink] = React.useState<boolean>(false);
  const [state, setState] = React.useState<IUserLink>({
    name: '',
    url: '',
    description: ''
  });

  const { getLinks, updateLinks, deleteLink, links } = useLinks({ context });
  const { getTrending, trending } = useTrending({ context });

  const calcDate: (date: Date, weekDay?: boolean) => string = (date, weekDay = false) => {
    return date.toLocaleDateString('en-US', {
      year: 'numeric',
      day: 'numeric',
      month: 'long',
      weekday: weekDay ? undefined : 'long'
    });
  }
  

  const handleChange: (e: React.ChangeEvent<HTMLInputElement>, newValue?: string) => void = React.useCallback((e, newValue?) => {
    setState((s: IUserLink) => {
      return {
        ...s,
        [e.target.id]: newValue
      }
    });
  }, []);

  const toggleBack: () => void = React.useCallback(() => {
    const newDate: Date = new Date(date);
    newDate.setDate(date.getDate() - 1);
    setDate(newDate);
  }, [date]);

  const toggleForward: () => void = React.useCallback(() => {
    const newDate: Date = new Date(date);
    newDate.setDate(date.getDate() + 1);
    setDate(newDate);
  }, [date]);

  const openCreateLink: () => void = React.useCallback(() => {
    setCreateLink(true);
  }, []);
  
  const closeCreateLink: () => void = React.useCallback(() => {
    setCreateLink(false);
  }, []);

  const submitNewLink: () => Promise<void> = React.useCallback(async () => {
    await updateLinks(state);
    setCreateLink(false);
  }, [state, updateLinks]);

  const handleLink: (e: React.MouseEvent<HTMLDivElement>) => void = React.useCallback((e) => {
    window.open(e.currentTarget.id, '_blank noreferrer');
  }, []);

  const handleDelete: (e: React.MouseEvent<HTMLButtonElement>) => Promise<void> = React.useCallback(async (e) => {
    const link: IUserLink | undefined = links.find((l: IUserLink) => l.name === e.currentTarget.id);

    if (link) {
      Promise.resolve(deleteLink(link)).catch((e: unknown) => {
        console.log(e);
        setCreateLink(false);
      });
    }
  }, [links, deleteLink])

  React.useEffect(() => {
    Promise.resolve(getLinks()).catch((e: unknown) => console.log(e));

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  React.useEffect(() => {
    Promise.resolve(getTrending()).catch((e: unknown) => console.log(e));

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  
  return (
    <section className={styles.personalDashboard}>
      <div className={styles.dashboard}>
        <div className={styles.dashboardItem}>
          <Stack tokens={{ childrenGap: 10 }} style={{ width: '100%', position: 'relative' }}>
            <Text className={styles.sectionTitle}>
                Agenda for {calcDate(date)}
            </Text>
            <div className={styles.bodyHeader} style={{ padding: '0 12px' }}>
              <Text className={styles.headerText}>
                My Agenda
              </Text>
              <PrimaryButton iconProps={{ iconName: 'Back' }} onClick={toggleBack} className={styles.bodyButton} style={{ marginRight: '4px', width: '100px' }}>
                Back
              </PrimaryButton>
              <DefaultButton iconProps={{ iconName: 'Forward' }} onClick={toggleForward} className={styles.bodyButton} style={{ width: '100px', background: 'white' }} styles={{ flexContainer: { flexDirection: 'row-reverse' }}}>
                Forward
              </DefaultButton>              
           </div>
            <Agenda date={calcDate(date)} days={1}></Agenda>  
          </Stack>
        </div>
        <div className={styles.dashboardItem}>
          <Stack tokens={{ childrenGap: 10 }} style={{ width: '100%' }}>
              <Text className={styles.sectionTitle}>My Tasks</Text>
              <Tasks />
              <Todo />
          </Stack>
        </div>
        <div className={styles.dashboardItem}>
          <Stack tokens={{ childrenGap: 10 }} style={{ width: '100%' }}>
            <Text className={styles.sectionTitle}>
                My Links
            </Text>
            <div className={styles.bodyHeader} style={{ padding: '0 12px' }}>
              <Text className={styles.headerText}>
                Links
              </Text>
              <PrimaryButton
                className={styles.bodyButton}
                iconProps={{ iconName: 'Add' }}
                style={{ display: createLink ? 'none' : 'flex' }}
                onClick={openCreateLink}>
                Add
              </PrimaryButton>
            </div>
            <Stack role="ul" tokens={{ childrenGap: 4 }} className={styles.links} style={{ padding: '0 12px' }}>
                {createLink ?
                  <>
                      <Label>Add a link</Label>
                      <TextField placeholder='Enter link name' id='name' onChange={handleChange} />
                      <TextField placeholder='Enter link URL' id='url' onChange={handleChange} />
                      <TextField placeholder='Enter link description' id='description' multiline onChange={handleChange} />
                      <Stack horizontal tokens={{ childrenGap: 15 }} horizontalAlign='end'>
                        <PrimaryButton iconProps={{ iconName: 'Add' }} onClick={submitNewLink}>
                          Add
                        </PrimaryButton>
                        <DefaultButton iconProps={{ iconName: 'Cancel' }} style={{ background: 'white' }} onClick={closeCreateLink}>
                          Cancel
                        </DefaultButton>
                      </Stack>
                  </>
                :
                  links.map(link => (
                    <StackItem role="li" className={styles.linkItem} key={link.url}>
                      <TooltipHost content={link.description} styles={{ root: { display: 'flex', flexGrow: 1 }}}>
                        <PrimaryButton
                          href={link.url}
                          style={{ flexGrow: 1, textAlign: 'center', marginRight: '8px' }}
                          target='_blank noreferrer'>
                            {link.name}
                        </PrimaryButton>
                      </TooltipHost>
                      <PrimaryButton iconProps={{ iconName: 'Delete' }} id={link.name} onClick={handleDelete} />
                    </StackItem>                  
                  ))
                }
            </Stack>
          </Stack>
        </div>
        <div className={styles.dashboardItem}>
          <Stack tokens={{ childrenGap: 10 }} style={{ width: '100%' }}>
            <Text className={styles.sectionTitle}>
                My Trending Items
            </Text>
            <div className={styles.bodyHeader} style={{ padding: '0 12px' }}>
              <Text className={styles.headerText}>
                Trending
              </Text>
            </div>
            <Stack role="ul" tokens={{ childrenGap: 10 }} className={styles.links} style={{ padding: '0 12px' }}>
                {trending.map((trend: IUserTrending) => (
                  <StackItem role="li" className={styles.linkItem} style={{ cursor: 'pointer' }} key={trend.resourceReference.id}>
                    <File
                      fileQuery={trend.resourceReference.id}
                      id={trend.resourceReference.webUrl}
                      onClick={handleLink}
                    />
                  </StackItem>
                ))}
            </Stack>
          </Stack>
        </div>
      </div>
    </section>
  )
}

export default PersonalDashboard;
