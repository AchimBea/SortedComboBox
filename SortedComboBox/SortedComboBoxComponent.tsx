import * as React from "react";
import { ComboBox, IComboBoxOption, Icon } from "@fluentui/react";

export interface ISortedComboBoxComponentProps {
	label: string;
	value: number | null;
	options: ComponentFramework.PropertyHelper.OptionMetadata[];
	configuration: string | null;
	onChange: (newValue: number | undefined) => void;
	disabled: boolean;
	masked: boolean;
	formFactor: "small" | "large";
}

const iconStyles = { marginRight: "8px" };

const onRenderOption = (option?: IComboBoxOption): JSX.Element => {
	if (option) {
		return (
			<div>
				{option.data && option.data.icon && (
					<Icon style={iconStyles} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
				)}
				<span>{option.text}</span>
			</div>
		);
	}
	return <></>;
};

export const SortedComboBoxComponent = React.memo(
	(props: ISortedComboBoxComponentProps) => {

		const { label, value, options, onChange, configuration, disabled, masked, formFactor } = props;
		const valueKey = value != null ? value.toString() : undefined;

		const items = React.useMemo(() => {
			let iconMapping: Record<number, string> = {};
			let configError: string | undefined;
			if (configuration) {
				try {
					iconMapping = JSON.parse(configuration) as Record<number, string>;
				} catch {
					configError = `Invalid configuration: '${configuration}'`;
				}
			}

			return {
				error: configError,
				options: options.map((item) => {
					return {
						key: item.Value.toString(),
						data: { value: item.Value, icon: iconMapping[item.Value] },
						text: item.Label,
					} as IComboBoxOption;
				}),
			};
		}, [options, configuration]);

		const onChangeDropDown = React.useCallback(
			(ev: unknown, option?: IComboBoxOption): void => {
				onChange(option ? (option.data.value as number) : undefined);
			},
			[onChange]
		);

		return (
			<>
				{items.error}

				{masked && "****"}

				<ComboBox
					placeholder={"---"}
					label={label}
					ariaLabel={label}
					options={items.options}
					selectedKey={valueKey}
					disabled={disabled}
					onRenderOption={onRenderOption}
					onChange={onChangeDropDown}
				/>

			</>
		);
	}
);
SortedComboBoxComponent.displayName = "SortedComboBoxComponent";
