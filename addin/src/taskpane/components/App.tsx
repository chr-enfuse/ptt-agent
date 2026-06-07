import * as React from "react";
import Chat from "./Chat";
import TextInsertion from "./TextInsertion";
import { makeStyles, tokens, Text, Divider } from "@fluentui/react-components";
import { insertText } from "../taskpane";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
  },
  header: {
    padding: "8px 12px",
    backgroundColor: tokens.colorBrandBackground,
  },
  headerText: {
    color: tokens.colorNeutralForegroundOnBrand,
  },
  sanityCheck: {
    flexShrink: 0,
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <header className={styles.header}>
        <Text weight="semibold" size={400} className={styles.headerText}>
          {props.title}
        </Text>
      </header>
      <Chat />
      <Divider>Office.js sanity check (milestone 1)</Divider>
      <div className={styles.sanityCheck}>
        <TextInsertion insertText={insertText} />
      </div>
    </div>
  );
};

export default App;
