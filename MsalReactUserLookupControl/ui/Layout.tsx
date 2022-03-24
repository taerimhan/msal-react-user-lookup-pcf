import { mergeStyleSets } from "@fluentui/react/lib/Styling";
import * as React from "react";
import { Login } from "../components/Login";

const classNames = mergeStyleSets({
    container: {
        padding: "8px 16px 16px 16px",
        border: "1px solid #edebe9",
        margin: "8px",
        borderRadius: "2px",
        boxShadow: "0 0.2rem 0.3rem -0.075rem rgb(0 0 0 / 10%)"
    }
});

export interface ILayoutProps {
    children: any
}

export const Layout: React.FC<ILayoutProps> = (props: ILayoutProps) => {
    return (
        <div className={classNames.container}>
            <header>
                <Login />
            </header>
            <main>
                {props.children}
            </main>
        </div>
    );
}